Attribute VB_Name = "modSystemMonitorSupport"
Option Explicit

'//==================================================================
'//     Type Structure to store the System Values
'//==================================================================
Public Type SystemMonitorStruct
    CPULoadPercent                  As Single
    MemoryLoadPercent               As Single
    PhysicalMemoryTotal             As Long
    PhysicalMemoryAvailable         As Long
    PhysicalMemoryAvailablePercent  As Single
    PageFileTotal                   As Long
    PageFileAvailable               As Long
    PageFileAvailablePercent        As Single
    VirtualMemoryTotal              As Long
    VirtualMemoryAvailable          As Long
    VirtualMemoryAvailablePercent   As Single
    HDTotalBytes                    As Currency
    HDTotalFreeBytes                As Currency
    HDAvailableFreeBytes            As Currency
    HDTotalBytesUsed                As Currency
    HDAvailablePercent              As Single
End Type
'//==================================================================


'//==================================================================
'//     CPU Monitor CODE WinNT START
'//     ORIGINAL CODER: The KPD-Team at http://www.allapi.net
'//==================================================================
Private Const SYSTEM_BASICINFORMATION = 0&
Private Const SYSTEM_PERFORMANCEINFORMATION = 2&
Private Const SYSTEM_TIMEINFORMATION = 3&
Private Const NO_ERROR = 0

'// We use Currency instead of LARGE_INTEGER
'// So we don't have to convert back and forth
'Private Type LARGE_INTEGER
'    dwLow As Long
'    dwHigh As Long
'End Type
'
Private Type SYSTEM_BASIC_INFORMATION
    dwUnknown1 As Long
    uKeMaximumIncrement As Long
    uPageSize As Long
    uMmNumberOfPhysicalPages As Long
    uMmLowestPhysicalPage As Long
    uMmHighestPhysicalPage As Long
    uAllocationGranularity As Long
    pLowestUserAddress As Long
    pMmHighestUserAddress As Long
    uKeActiveProcessors As Long
    bKeNumberProcessors As Byte
    bUnknown2 As Byte
    wUnknown3 As Integer
End Type

Private Type SYSTEM_PERFORMANCE_INFORMATION
    liIdleTime As Currency
    dwSpare(0 To 75) As Long
End Type

Private Type SYSTEM_TIME_INFORMATION
    liKeBootTime As Currency
    liKeSystemTime As Currency
    liExpTimeZoneBias  As Currency
    uCurrentTimeZoneId As Long
    dwReserved As Long
End Type

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function NtQuerySystemInformation Lib "ntdll" (ByVal dwInfoType As Long, ByVal lpStructure As Long, ByVal dwSize As Long, ByVal dwReserved As Long) As Long

Private liOldIdleTime       As Currency
Private liOldSystemTime     As Currency
'//==================================================================
'//     CPU Monitor CODE WinNT END
'//==================================================================


'//==================================================================
'//     CPU Monitor CODE Win9X START
'//     ORIGINAL CODER: The KPD-Team at http://www.allapi.net
'//==================================================================
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const HKEY_DYN_DATA = &H80000006
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const ERROR_SUCCESS = 0&
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private hKey As Long, dwDataSize As Long, dwCpuUsage As Byte, dwType As Long
'//==================================================================
'//     CPU Monitor CODE Win9X END
'//==================================================================


'//==================================================================
'//     HARD DRIVE DECLARES
'//==================================================================
Private Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Private Declare Function GetDiskFreeSpace Lib "Kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "Kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "Kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

'NOTE!!
'GetDiskFreeSpaceEx function return value definition
'lpFreeBytesAvailable
'  [out] Pointer to a variable that receives the total number of free bytes on
'  the disk that are available to the user associated with the calling thread.
'  Windows 2000: If per-user quotas are in use, this value may be less than the total number of free bytes on the disk.
'
'lpTotalNumberOfFreeBytes
'  [out] Pointer to a variable that receives the total number of free bytes on the disk.
'  This parameter can be NULL.
'//==================================================================


'//==================================================================
'//     MEMORY MONITOR CODE START
'//     ORIGINAL CODER: Randy Birch at VBnet (http://www.mvps.org/vbnet/index.html)
''//==================================================================
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailablePageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Declare Sub GlobalMemoryStatus Lib "Kernel32" (lpBuffer As MEMORYSTATUS)
'//==================================================================
'//     MEMORY MONITOR CODE END
'//==================================================================



'// Return True is the OS is WindowsNT3.5(1), NT4.0, 2000, Me, XP
Public Function IsWinNTInstalled() As Boolean
    Dim OSInfo As OSVERSIONINFO
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    'retrieve OS version info
    GetVersionEx OSInfo
    'if we're on NT, return True
    IsWinNTInstalled = (OSInfo.dwPlatformId = 2)
End Function


'// Setup the nesaccary options to start retreiving CPU Permormance data
Public Sub CPUInitialize(ByRef bWindowsNT As Boolean)
Dim SysTimeInfo     As SYSTEM_TIME_INFORMATION
Dim SysPerfInfo     As SYSTEM_PERFORMANCE_INFORMATION
Dim ret             As Long
    
    '// Initialize the CPU Load settings based on if user is using Windows NT or 9x
    If bWindowsNT Then
        '// get new system time
        ret = NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, VarPtr(SysTimeInfo), LenB(SysTimeInfo), 0&)
        If ret <> NO_ERROR Then
            Debug.Print "Error while initializing the system's time!", vbCritical
            Exit Sub
        End If
        
        '// get new CPU's idle time
        ret = NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, VarPtr(SysPerfInfo), LenB(SysPerfInfo), ByVal 0&)
        If ret <> NO_ERROR Then
            Debug.Print "Error while initializing the CPU's idle time!", vbCritical
            Exit Sub
        End If
        
        '// store new CPU's idle and system time
        liOldIdleTime = SysPerfInfo.liIdleTime
        liOldSystemTime = SysTimeInfo.liKeSystemTime
    
    Else
        '// start the counter by reading the value of the 'StartStat' key
        If RegOpenKeyEx(HKEY_DYN_DATA, "PerfStats\StartStat", 0, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then
            Debug.Print "Error while initializing counter"
            Exit Sub
        End If
        dwDataSize = 4 'Length of Long
        RegQueryValueEx hKey, "KERNEL\CPUUsage", ByVal 0&, dwType, dwCpuUsage, dwDataSize
        RegCloseKey hKey
        
        '// get current counter's value
        If RegOpenKeyEx(HKEY_DYN_DATA, "PerfStats\StatData", 0, KEY_READ, hKey) <> ERROR_SUCCESS Then
            Debug.Print "Error while opening counter key"
            Exit Sub
        End If
    
    End If
    
    
End Sub


'// Return the current CPU Load based on if the user is running Windows 9x or NT
Public Function CPUQuery(ByRef bWindowsNT As Boolean) As Single
Dim SysBaseInfo     As SYSTEM_BASIC_INFORMATION
Dim SysPerfInfo     As SYSTEM_PERFORMANCE_INFORMATION
Dim SysTimeInfo     As SYSTEM_TIME_INFORMATION
Dim dbIdleTime      As Currency
Dim dbSystemTime    As Currency
Dim ret             As Long
    
    If bWindowsNT Then
        '// Set Initial Value
        CPUQuery = -1
        
        '// get number of processors in the system
        ret = NtQuerySystemInformation(SYSTEM_BASICINFORMATION, VarPtr(SysBaseInfo), LenB(SysBaseInfo), 0&)
        If ret <> NO_ERROR Then
            Debug.Print "Error while retrieving the number of processors!", vbCritical
            Exit Function
        End If
        
        '// get new system time
        ret = NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, VarPtr(SysTimeInfo), LenB(SysTimeInfo), 0&)
        If ret <> NO_ERROR Then
            Debug.Print "Error while retrieving the system's time!", vbCritical
            Exit Function
        End If
        
        '// get new CPU's idle time
        ret = NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, VarPtr(SysPerfInfo), LenB(SysPerfInfo), ByVal 0&)
        If ret <> NO_ERROR Then
            Debug.Print "Error while retrieving the CPU's idle time!", vbCritical
            Exit Function
        End If
        
        '// CurrentValue = NewValue - OldValue
        dbIdleTime = (SysPerfInfo.liIdleTime) - (liOldIdleTime)
        dbSystemTime = (SysTimeInfo.liKeSystemTime) - (liOldSystemTime)
        
        '// CurrentCpuIdle = IdleTime / SystemTime
        If dbSystemTime <> 0 Then dbIdleTime = dbIdleTime / dbSystemTime
        
        '// CurrentCpuUsage% = 100 - (CurrentCpuIdle * 100) / NumberOfProcessors
        dbIdleTime = 100 - dbIdleTime * 100 / SysBaseInfo.bKeNumberProcessors + 0.5
        
        '// Return Query Value
        CPUQuery = CLng(dbIdleTime)
        
        '// store new CPU's idle and system time
        liOldIdleTime = SysPerfInfo.liIdleTime
        liOldSystemTime = SysTimeInfo.liKeSystemTime
    
    Else
    
        '// Size of Long
        dwDataSize = 4
        
        '// Query the counter in the Registry
        If RegQueryValueEx(hKey, "KERNEL\CPUUsage", ByVal 0&, dwType, dwCpuUsage, dwDataSize) = NO_ERROR Then
            '// Return Value
            CPUQuery = CLng(dwCpuUsage)
        Else
            MsgBox "Unable to Query Registry!"
            CPUQuery = -1
        End If
        
    End If
    
    
End Function


'// Stop the CPU Query
Public Sub CPUTerminate(ByRef bWindowsNT As Boolean)

    '// We only need to do this if NOT WindowsNT.
    '// For WindowsNT based systems there is nothing we need to do
    If Not bWindowsNT Then
        '// Close the Registry
        RegCloseKey hKey
        '// Stop the counter
        If RegOpenKeyEx(HKEY_DYN_DATA, "PerfStats\StopStat", 0, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then
            Debug.Print "Error while stopping counter"
            Exit Sub
        End If
        dwDataSize = 4 'length of Long
        RegQueryValueEx hKey, "KERNEL\CPUUsage", ByVal 0&, dwType, dwCpuUsage, dwDataSize
        RegCloseKey hKey
    End If
    
    
End Sub


'// Get the current MEMORY levels
Public Sub MEMORYQuery(oSystemMonitor As SystemMonitorStruct)
Dim tMem As MEMORYSTATUS

    GlobalMemoryStatus tMem
    
    '// Populate SystemMonitorSturct with the MEMORY Details
    oSystemMonitor.MemoryLoadPercent = tMem.dwMemoryLoad
    oSystemMonitor.PhysicalMemoryTotal = tMem.dwTotalPhys
    oSystemMonitor.PhysicalMemoryAvailable = tMem.dwAvailPhys
    oSystemMonitor.PageFileTotal = tMem.dwTotalPageFile
    oSystemMonitor.PageFileAvailable = tMem.dwAvailablePageFile
    oSystemMonitor.VirtualMemoryTotal = tMem.dwTotalVirtual
    oSystemMonitor.VirtualMemoryAvailable = tMem.dwAvailVirtual
    
    '// Calculate the Percentages
    oSystemMonitor.PhysicalMemoryAvailablePercent = (1 - (tMem.dwAvailPhys / tMem.dwTotalPhys)) * 100
    oSystemMonitor.PageFileAvailablePercent = (1 - (tMem.dwAvailablePageFile / tMem.dwTotalPageFile)) * 100
    oSystemMonitor.VirtualMemoryAvailablePercent = (1 - (tMem.dwAvailVirtual / tMem.dwTotalVirtual)) * 100
    
End Sub


'// Get the HD total space, free space, used space
Public Function HDQuery(ByVal strDrive As String, oSystemMonitor As SystemMonitorStruct)
Dim curAvailableBytesFree   As Currency
Dim curTotalBytes           As Currency
Dim curTotalBytesFree       As Currency
Dim curTotalBytesUsed       As Currency
Dim curAvailablePercent     As Currency

Dim lngSectors              As Long
Dim lngBytesPerSector       As Long
Dim lngFreeClusters         As Long
Dim lngTotalClusters        As Long
Dim lngDrvSpaceTotal        As Long
Dim lngDrvSpaceFree         As Long

Dim lngPtr                  As Long
  
    '// See if we can get a pointer to the GetDiskFreeSpaceExA API in the kernel32.dll file
    '// If we can then we know we can use the GetDiskFreeSpaceEx function.
    lngPtr = GetProcAddress(GetModuleHandle("kernel32.dll"), "GetDiskFreeSpaceExA")
    
    If lngPtr <> 0 Then
        '// Get drive info using GetDiskFreeSpaceEx
        If GetDiskFreeSpaceEx(strDrive, curAvailableBytesFree, curTotalBytes, curTotalBytesFree) <> 0 Then
            '// adjust by multiplying the returned value by 10,000 to remove
            '// the decimal places the currency data type returns.
            curTotalBytes = curTotalBytes * 10000
            curTotalBytesFree = curTotalBytesFree * 10000
            curAvailableBytesFree = curAvailableBytesFree * 10000
            curTotalBytesUsed = CCur((curTotalBytes - curAvailableBytesFree))
            curAvailablePercent = CSng((curAvailableBytesFree / curTotalBytes) * 100)
        End If
        
    Else
        '// Get drive info using GetDiskFreeSpace
        If GetDiskFreeSpace(strDrive, lngSectors, lngBytesPerSector, lngFreeClusters, lngTotalClusters) <> 0 Then
            '// Calculate to get the data
            On Local Error Resume Next
            curTotalBytes = CCur(lngSectors) * CCur(lngBytesPerSector) * CCur(lngTotalClusters)
            curTotalBytesFree = CCur(lngSectors) * CCur(lngBytesPerSector) * CCur(lngFreeClusters)
            curTotalBytesUsed = CCur(curTotalBytes - curTotalBytesFree)
            curAvailablePercent = CSng((curTotalBytesFree / curTotalBytes) * 100)
            curAvailableBytesFree = curTotalBytesFree
            On Local Error GoTo 0
        End If
        
    End If
    
    '// Return Values
    oSystemMonitor.HDTotalBytes = curTotalBytes
    oSystemMonitor.HDTotalFreeBytes = curTotalBytesFree
    oSystemMonitor.HDTotalBytesUsed = curTotalBytesUsed
    oSystemMonitor.HDAvailableFreeBytes = curAvailableBytesFree
    oSystemMonitor.HDAvailablePercent = curAvailablePercent
    
End Function


Public Function DetermineWindowsFolder() As String
Dim strWinDir   As String
Dim lngRet      As Long
    
    strWinDir = String(255&, 0)
    lngRet = GetWindowsDirectory(strWinDir, 255&)
    DetermineWindowsFolder = Left(strWinDir, lngRet)

End Function


