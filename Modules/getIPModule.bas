Attribute VB_Name = "getIPModule"
'==     ȡ����ʱIP���࣬ԭ��ͨ��������ȡ�÷������ϴ������IP��ע����ֵ   ==
    
  '==   ������VB�����һ����ģ�飬Ҳ����ֱ�ӷ����������ڵĴ����ʹ��ʱ����   GetIP   ���ɣ�   ==
  Option Explicit
    
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
  '����Զ��ע���
Public Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
    
Public Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Public Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
    
  Const HKEY_LOCAL_MACHINE = &H80000002
  Const ERROR_SUCCESS = 0&
  Const REG_SZ = 1
  Const ERROR_NONE = 0
  Const REG_DWORD = 4
  Const REG_MULTI_SZ = 7
  Const REG_STRING = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces"
  Const REG_STRING1 = "SYSTEM\CurrentControlSet\Services"
  '   ע���ؼ��ְ�ȫѡ��...
  Const READ_CONTROL = &H20000
  Const KEY_QUERY_VALUE = &H1
  Const KEY_SET_VALUE = &H2
  Const KEY_CREATE_SUB_KEY = &H4
  Const KEY_ENUMERATE_SUB_KEYS = &H8
  Const KEY_NOTIFY = &H10
  Const KEY_CREATE_LINK = &H20
  Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                                                KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                                                KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
    
  'ȡ������IP
  Function GetIP() As String
        Dim tstr1     As String
        EnumKey HKEY_LOCAL_MACHINE, REG_STRING, 1, tstr1
        tstr1 = QueryValue(HKEY_LOCAL_MACHINE, REG_STRING & _
                                        "\" & tstr1, "DefaultGateway")
        GetIP = GetRegValue(tstr1)
  End Function
    
Public Function GetRegValue(ByVal ServerName As String) As String
        Dim tstr1     As String
        Dim lngReg     As Long
          
        RegConnectRegistry ServerName, HKEY_LOCAL_MACHINE, lngReg
          
        EnumKey lngReg, REG_STRING, 0, tstr1
          
        GetRegValue = QueryValue(lngReg, REG_STRING & "\" & tstr1, "DhcpIPAddress")
  End Function
    
Public Function EnumKey(hMainKey As Long, sSubKey As String, lIndex As Long, lpStr As String) As Boolean
  'EnumKey��������hMainKey������sSubKey�Ӽ�ָ����ע�����lIndexΪҪ��ѯ���Ӽ�ֵ
  '������,lpStrΪ�����Ӽ�ֵ���ַ������壬���Ҫ��ѯһ����ֵ�������Ӽ���ֻҪ��lIndex
  '��������Ϊ0��Ȼ��lIndex����1�ٵ���EnumKey������ֱ����������0Ϊֹ
        Dim hKey     As Long           '�򿪼��ľ��
        Dim i     As Long
          
        If RegOpenKey(hMainKey, sSubKey, hKey) = ERROR_SUCCESS Then
              lpStr = Space(255) + Chr(0)
              If RegEnumKey(hKey, lIndex, lpStr, Len(lpStr)) = ERROR_SUCCESS Then
                    EnumKey = True
              Else
                    EnumKey = False
              End If
        Else
              EnumKey = False
        End If
        RegCloseKey hKey
  End Function
    
Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
        Dim lRetVal     As Long
        Dim hKey     As Long                     '�򿪼��ľ��
        Dim vValue     As Variant
          
        lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
        lRetVal = QueryValueEx(hKey, sValueName, vValue)
        QueryValue = vValue
        RegCloseKey (hKey)
  End Function
    
Public Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
        Dim cch     As Long
        Dim lrc     As Long
        Dim lType     As Long
        Dim lValue     As Long
        Dim sValue     As String
          
        On Error GoTo QueryValueExError
          
        lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
        If lrc <> ERROR_NONE Then Error 5
          
        Select Case lType
              '��ѯ�ַ���ֵ
              Case REG_SZ:
                    sValue = String(cch, 0)
                    lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
                    If lrc = ERROR_NONE Then
                          vValue = Left$(sValue, cch)
                    Else
                          vValue = Empty
                    End If
                
              '��ѯ����ֵ
              Case REG_DWORD:
                    lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
                    If lrc = ERROR_NONE Then vValue = lValue
              Case REG_MULTI_SZ:
                    sValue = String(cch, 0)
                    lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
                    If lrc = ERROR_NONE Then
                          vValue = Left$(sValue, cch)
                    Else
                          vValue = Empty
                    End If
              Case Else
                  lrc = -1
        End Select
          
QueryValueExExit:
          
        QueryValueEx = lrc
        Exit Function
          
QueryValueExError:
          
        Resume QueryValueExExit
    
  End Function

