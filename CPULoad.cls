VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPULoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const ClassName As String = "CPULoad"

Private Const Err_Initialize As Long = vbObjectError + 8001
Private Const Err_UnableToStartPerfmon As Long = vbObjectError + 8002
Private Const Err_CPUIndexOOB As Long = vbObjectError + 8003
Private Const Err_CantFindProcessorPerfMon As Long = vbObjectError + 8004
Private Const Err_CantFindCPUUsagePerfMon As Long = vbObjectError + 8005
Private Const Err_UnableToReadPDB As Long = vbObjectError + 8006

Private Declare Sub Memcopy Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SystemTime, lpFileTime As Currency) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

Private Const HKEY_DYN_DATA = &H80000006
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const REG_DWORD = 4
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_MORE_DATA = 234

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0

Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000

Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or _
                                KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or _
                                KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or _
                                KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or _
                        KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SystemTime
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Type PERF_INSTANCE_DEFINITION
    ByteLength As Long
    ParentObjectTitleIndex As Long
    ParentObjectInstance As Long
    UniqueID As Long
    NameOffset As Long
    NameLength As Long
End Type

Private Type PERF_COUNTER_BLOCK
    ByteLength As Long
End Type

Private Type PERF_DATA_BLOCK
    Signature As String * 4
    LittleEndian As Long
    Version As Long
    Revision As Long
    TotalByteLength As Long
    HeaderLength As Long
    NumObjectTypes As Long
    DefaultObject As Long
    SystemTime As SystemTime
    PerfTime As LARGE_INTEGER
    PerfFreq As LARGE_INTEGER
    PerTime100nSec As LARGE_INTEGER
    SystemNameLength As Long
    SystemNameOffset As Long
End Type

Private Type PERF_OBJECT_TYPE
    TotalByteLength As Long
    DefinitionLength As Long
    HeaderLength As Long
    ObjectNameTitleIndex As Long
    ObjectNameTitle As Long
    ObjectHelpTitleIndex As Long
    ObjectHelpTitle As Long
    DetailLevel As Long
    NumCounters As Long
    DefaultCounter As Long
    NumInstances As Long
    CodePage As Long
    PerfTime As LARGE_INTEGER
    PerfFreq As LARGE_INTEGER
End Type

Private Type PERF_COUNTER_DEFINITION
    ByteLength As Long
    CounterNameTitleIndex As Long
    CounterNameTitle As Long
    CounterHelpTitleIndex As Long
    CounterHelpTitle As Long
    DefaultScale As Long
    DetailLevel As Long
    CounterType As Long
    CounterSize As Long
    CounterOffset As Long
End Type


'-------------------------------------------------------------------
Private Const Processor_IDX_Str As String = "238"
Private Const Processor_IDX  As Long = 238
Private Const CPUUsageIDX As Long = 6

Private m_lProcessorsCount As Long
Private m_lBufferSize As Long
Private m_bIsWinNT As Boolean

Private m_bW9xCollecting As Boolean
Private m_lW9xCpuUsage As Long
Private m_hW9xCpuKey As Long

Private PDB As PERF_DATA_BLOCK
Private POT As PERF_OBJECT_TYPE
Private PCD As PERF_COUNTER_DEFINITION
Private PID As PERF_INSTANCE_DEFINITION
Private PCB As PERF_COUNTER_BLOCK

Private VI As OSVERSIONINFO

Private SysTime As Currency
Private PrevSysTime As Currency
Private m_aCounters() As Currency
Private m_aPrevCounters() As Currency

Private Const ByteIncrement As Long = 4096

Private Sub Class_Initialize()
    
    VI.dwOSVersionInfoSize = Len(VI)
    
    If GetVersionEx(VI) = 0 Then
        Err.Raise Err_Initialize, ClassName & ".Initialize", "Can't get the Windows version"
    End If

    m_bIsWinNT = (VI.dwPlatformId = VER_PLATFORM_WIN32_NT)
    m_lProcessorsCount = -1
    m_lBufferSize = ByteIncrement
    
End Sub

Private Sub Class_Terminate()
    ReleaseCPUData
End Sub

Public Function CollectCPUData() As Boolean
Dim H As Long, R As Long
Dim aBuf() As Byte, lAllocSz As Long
Dim lSrc As Long, lDest As Long
Dim ptrPOT As Long, ptrPCB As Long
Dim i As Long, lCPU As Long
Dim ST As Currency
Dim sInstanceName As String

    If m_bIsWinNT = True Then
        '// For WinNT Systems
        
        '// Allocate the buffer.
        lAllocSz = m_lBufferSize    '// Initial allocation size
        ReDim aBuf(1 To lAllocSz) As Byte
        
        '// We loop until RQVex says that our buffer is large enough
        While RegQueryValueEx(HKEY_PERFORMANCE_DATA, Processor_IDX_Str, _
                              0&, 0&, aBuf(1), m_lBufferSize) = ERROR_MORE_DATA
            '// Get a Buffer that is big enough.
            '// Increment the allocation size
            lAllocSz = lAllocSz + ByteIncrement
            ReDim aBuf(1 To lAllocSz) As Byte
            '// Tell RQVex how big we allocated the buffer
            m_lBufferSize = lAllocSz
        Wend
        
        lDest = VarPtr(PDB)
        lSrc = VarPtr(aBuf(1))
        Memcopy ByVal lDest, ByVal lSrc, LenB(PDB)
        '// Because RegQueryValueEx modifies the
        '// Data in BufferSize, reset it to the
        '// Proper value for the buffer size
        '// We want to save the size so that next time we hopefully
        '// wont have to loop so much to find the size needed
        m_lBufferSize = lAllocSz
        
        '// Check for success
        If PDB.Signature <> "PERF" Then
            Err.Raise Err_UnableToReadPDB, ClassName & ".CollectCPUData()", "Unable to read performance data"
        End If
        
        '==========================================================================
        '// Locate the performance object
        lDest = VarPtr(POT)
        lSrc = VarPtr(aBuf(1)) + PDB.HeaderLength
        For i = 1 To PDB.NumObjectTypes
            Memcopy ByVal lDest, ByVal lSrc, LenB(POT)
            ptrPOT = lSrc
            If POT.ObjectNameTitleIndex = Processor_IDX Then Exit For
            lSrc = lSrc + POT.TotalByteLength
        Next i
        
        '// Check for success
        If POT.ObjectNameTitleIndex <> Processor_IDX Then
            Err.Raise Err_CantFindProcessorPerfMon, ClassName & ".CollectData", "Unable to locate the 'Processor' performance object"
        End If
        
        '// Get the Processor Count
        If m_lProcessorsCount < 1 Then
            m_lProcessorsCount = GetCPUCount()
        End If
        
        '==========================================================================
        '// Locate the "% CPU usage" counter definition
        lDest = VarPtr(PCD)
        lSrc = lSrc + POT.HeaderLength
        For i = 1 To POT.NumCounters
            Memcopy ByVal lDest, ByVal lSrc, LenB(PCD)
            If PCD.CounterNameTitleIndex = CPUUsageIDX Then Exit For
            lSrc = lSrc + PCD.ByteLength
        Next i
        
        '// Check for success
        If PCD.CounterNameTitleIndex <> CPUUsageIDX Then
            Err.Raise Err_CantFindCPUUsagePerfMon, ClassName & ".CollectData", "Unable to locate the '% of CPU usage' performance counter"
        End If
        
        '==========================================================================
        '// Collecting counters
        lSrc = ptrPOT + POT.DefinitionLength
        For i = 1 To POT.NumInstances
            lDest = VarPtr(PID)
            Memcopy ByVal lDest, ByVal lSrc, LenB(PID)
            
            '// Get the Instance name
            '// The "-2" is because we dont need the terminating double null
            sInstanceName = Space(PID.NameLength - 2)
            Memcopy ByVal sInstanceName, ByVal lSrc + PID.NameOffset, PID.NameLength - 2
            sInstanceName = StrConv(sInstanceName, vbFromUnicode)
            
            lSrc = lSrc + PID.ByteLength
            lDest = VarPtr(PCB)
            Memcopy ByVal lDest, ByVal lSrc, LenB(PCB)
            ptrPCB = lSrc
            
            '// Win2K has an instance for the '% CPU usage' named '_Total' we dont want to
            '// report on that instance, we want the actual processor
            '// The processor instances are named 0,1,2, etc...
            
            'If sInstanceName <> "_Total" Then
            If IsNumeric(sInstanceName) Then
            
                '// Assumption.. the instance name will
                '// be an integer for the cpu index
                '// ie "0" = 0 = first cpu
                lCPU = CLng(sInstanceName)
                
                m_aPrevCounters(lCPU) = m_aCounters(lCPU)
                Memcopy ByVal VarPtr(m_aCounters(lCPU)), ByVal ptrPCB + PCD.CounterOffset, LenB(m_aCounters(lCPU))
            End If
            
            lSrc = lSrc + PCB.ByteLength
        Next i
        
        '==========================================================================
        PrevSysTime = SysTime
        SystemTimeToFileTime PDB.SystemTime, ST
        SysTime = ST

    Else
        '// For Win9x Systems
        If Not m_bW9xCollecting Then
            R = RegOpenKeyEx(HKEY_DYN_DATA, "PerfStats\StartStat", 0&, KEY_ALL_ACCESS, H)
            If R <> ERROR_SUCCESS Then
                Err.Raise Err_UnableToStartPerfmon, ClassName & ".CollectCPRData()", "Unable to start performance monitoring"
            End If
            
            Call RegQueryValueEx(H, "KERNEL\CPUUsage", 0&, REG_DWORD, m_lW9xCpuUsage, LenB(m_lW9xCpuUsage))
            Call RegCloseKey(H)
            
            R = RegOpenKeyEx(HKEY_DYN_DATA, "PerfStats\StatData", 0&, KEY_READ, m_hW9xCpuKey)
            If R <> ERROR_SUCCESS Then
                Err.Raise Err_UnableToReadPDB, ClassName & ".CollectCPUData()", "Unable to read performance data"
            End If
            
            m_bW9xCollecting = True
        End If
        
        Call RegQueryValueEx(m_hW9xCpuKey, "KERNEL\CPUUsage", 0&, REG_DWORD, m_lW9xCpuUsage, LenB(m_lW9xCpuUsage))
    End If
    
End Function

Public Function GetCPUCount() As Long
Dim SI As SYSTEM_INFO
    
    If m_lProcessorsCount < 1 Then
        GetSystemInfo SI
        GetCPUCount = SI.dwNumberOrfProcessors
        m_lProcessorsCount = SI.dwNumberOrfProcessors
        
        'Zero based array
        ReDim m_aPrevCounters(0 To m_lProcessorsCount - 1) As Currency
        ReDim m_aCounters(0 To m_lProcessorsCount - 1) As Currency
    Else
        GetCPUCount = m_lProcessorsCount
    End If
    
End Function

Public Function GetCPUUsage(Optional ByVal CPU_Index As Long = 1) As Long
'NOTE*** Our Counter Arrays are 0 Based, but what is passed is 1 based..
'Function Returns 0 to 100
    
    'deincrement for our internal array
    CPU_Index = CPU_Index - 1
    
    If m_bIsWinNT Then
        If m_lProcessorsCount < 0 Then CollectCPUData
        
        If (CPU_Index >= m_lProcessorsCount) Or (CPU_Index < 0) Then
            Err.Raise Err_CPUIndexOOB, ClassName & ".GetCPUUsage()", "CPU index out of bounds"
        End If
        
        If PrevSysTime = SysTime Then
            GetCPUUsage = 0
        Else
            GetCPUUsage = CLng(100 * (1 - (m_aCounters(CPU_Index) - m_aPrevCounters(CPU_Index)) / (SysTime - PrevSysTime)))
        End If
    Else
        If CPU_Index <> 0 Then
            Err.Raise Err_CPUIndexOOB, ClassName & ".GetCPUUsage()", "CPU index out of bounds"
        End If
        
        If Not m_bW9xCollecting Then CollectCPUData
        GetCPUUsage = m_lW9xCpuUsage
    End If
        
    'If GetCPUUsage < 0 Then GetCPUUsage = 0
    'If GetCPUUsage > 100 Then GetCPUUsage = 100
     
End Function

Private Sub ReleaseCPUData()
Dim H As Long
Dim R As Long

    If m_bIsWinNT Then Exit Sub
    If Not m_bW9xCollecting Then Exit Sub
    
    m_bW9xCollecting = False
    
    Call RegCloseKey(m_hW9xCpuKey)
    m_hW9xCpuKey = 0
    
    R = RegOpenKeyEx(HKEY_DYN_DATA, "PerfStats\StopStat", 0, KEY_ALL_ACCESS, H)
    If R <> ERROR_SUCCESS Then Exit Sub

    Call RegQueryValueEx(H, "KERNEL\CPUUsage", 0&, REG_DWORD, m_lW9xCpuUsage, LenB(m_lW9xCpuUsage))
    Call RegCloseKey(H)

End Sub

