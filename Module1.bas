Attribute VB_Name = "Module1"
Option Explicit
Type SYSTEM_INFO
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
Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type
Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Public Const PROCESSOR_INTEL_386 = 386
Public Const PROCESSOR_INTEL_486 = 486
Public Const PROCESSOR_INTEL_PENTIUM = 586
Public Const PROCESSOR_MIPS_R4000 = 4000
Public Const PROCESSOR_ALPHA_21064 = 21064

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As PointAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Const ERROR_SUCCESS As Long = 0
Public Const WS_VERSION_REQD As Long = &H101
Public Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD As Long = 1
Public Const SOCKET_ERROR As Long = -1
Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Public Type HOSTENT
  hName As Long
  hAliases As Long
  hAddrType As Integer
  hLen As Integer
  hAddrList As Long
End Type

Public Type WSADATA
  wVersion As Integer
  wHighVersion As Integer
  szDescription(0 To MAX_WSADescription) As Byte
  szSystemStatus(0 To MAX_WSASYSStatus) As Byte
  wMaxSockets As Integer
  wMaxUDPDG As Integer
  dwVendorInfo As Long
End Type



Public Type PointAPI
    x As Long
    y As Long
End Type
Public Const SWP_NOMOVE = &H2, SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1, HWND_NOTOPMOST = -2
Public Const RGN_DIFF = 4
Public Const SC_CLICKMOVE = &HF012&
Public Const WM_SYSCOMMAND = &H112

Dim CurRgn, TempRgn As Long

Public Function SetColorTransparent(bg As Form, transColor)
Dim x As Integer
Dim y As Integer
Dim Success As Long
CurRgn = CreateRectRgn(0, 0, bg.ScaleWidth, bg.ScaleHeight)

While y <= bg.ScaleHeight
    While x <= bg.ScaleWidth
        If GetPixel(bg.hdc, x, y) = transColor Then
            TempRgn = CreateRectRgn(x, y, x + 1, y + 1)
            Success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
        End If
        x = x + 1
    Wend
        y = y + 1
        x = 0
Wend
Success = SetWindowRgn(bg.hWnd, CurRgn, True)
DeleteObject (CurRgn)

End Function
Public Function GetIPAddress() As String
Dim sHostName As String * 256
Dim lpHost As Long
Dim HOST As HOSTENT
Dim dwIPAddr As Long
Dim tmpIPAddr() As Byte
Dim i As Integer
Dim sIPAddr As String

If Not SocketsInitialize() Then
  GetIPAddress = ""
  Exit Function
End If

If gethostname(sHostName, 256) = SOCKET_ERROR Then
  GetIPAddress = ""
  MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
  " has occurred. Unable to successfully get Host Name."
  SocketsCleanup
  Exit Function
End If

sHostName = Trim$(sHostName)
lpHost = gethostbyname(sHostName)

If lpHost = 0 Then
  GetIPAddress = ""
  MsgBox "Windows Sockets are not responding. " & _
  "Unable to successfully get Host Name."
  SocketsCleanup
  Exit Function
End If

CopyMemory HOST, lpHost, Len(HOST)
CopyMemory dwIPAddr, HOST.hAddrList, 4

ReDim tmpIPAddr(1 To HOST.hLen)

CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen

For i = 1 To HOST.hLen
sIPAddr = sIPAddr & tmpIPAddr(i) & "."
Next

GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
SocketsCleanup

End Function

Public Function GetIPHostName() As String

Dim sHostName As String * 256

If Not SocketsInitialize() Then
  GetIPHostName = ""
  Exit Function
End If

If gethostname(sHostName, 256) = SOCKET_ERROR Then
  GetIPHostName = ""
  MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
  " has occurred. Unable to successfully get Host Name."
  SocketsCleanup
  Exit Function
End If

GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
SocketsCleanup

End Function

Public Function HiByte(ByVal wParam As Integer)

HiByte = wParam \ &H100 And &HFF&

End Function

Public Function LoByte(ByVal wParam As Integer)
LoByte = wParam And &HFF&
End Function



Public Sub SocketsCleanup()

If WSACleanup() <> ERROR_SUCCESS Then
  MsgBox "Socket error occurred in Cleanup."
End If

End Sub



Public Function SocketsInitialize() As Boolean

Dim WSAD As WSADATA
Dim sLoByte As String
Dim sHiByte As String

If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
  MsgBox "The 32-bit Windows Socket is not responding."
  SocketsInitialize = False
  Exit Function
End If

If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
  MsgBox "This application requires a minimum of " & _
  CStr(MIN_SOCKETS_REQD) & " supported sockets."
  SocketsInitialize = False
  Exit Function
End If

If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
(LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
  sHiByte = CStr(HiByte(WSAD.wVersion))
  sLoByte = CStr(LoByte(WSAD.wVersion))
  MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
  " is not supported by 32-bit Windows Sockets."
  SocketsInitialize = False
  Exit Function
End If

'must be OK, so lets do it
SocketsInitialize = True

End Function

