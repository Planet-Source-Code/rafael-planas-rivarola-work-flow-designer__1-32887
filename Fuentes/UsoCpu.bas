Attribute VB_Name = "UsoCpu"
Type LARGE_INTEGER
                lowpart As Long
                highpart As Long
End Type

Declare Function QueryPerformanceCounter Lib _
        "kernel32" (lpPerformanceCount As LARGE_INTEGER) _
        As Long
Declare Function QueryPerformanceFrequency _
        Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long


Public Const REG_DWORD = 4 ' 32-bit number
Public Const HKEY_DYN_DATA = &H80000006

Declare Function RegQueryValueEx Lib _
        "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey _
        As Long, ByVal lpValueName As String, ByVal _
        lpReserved As Long, lpType As Long, lpData _
        As Any, lpcbData As Long) As Long

Declare Function RegOpenKey Lib "advapi32.dll" _
        Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal _
        lpSubKey As String, phkResult As Long) As Long

Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long


Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Global Qry As Long


Sub InitCPU()
    Exit Sub
    Dim lData As Long, lType As Long, lSize As Long
    Dim hKey As Long

    Qry = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StartStat", hKey)
    If Qry <> 0 Then
          MsgBox "No Puedo abrir las estadisticas de Uso de CPU"
          End
    End If
                
    lType = REG_DWORD
    lSize = 4
    Qry = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, lType, lData, lSize)
    Qry = RegCloseKey(hKey)

End Sub

