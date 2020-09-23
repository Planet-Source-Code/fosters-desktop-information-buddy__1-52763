VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3255
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   409
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   217
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1080
      Top             =   4260
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Sub ReleaseCapture Lib "user32" ()

Private Const CCM_FIRST = &H2000
Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Private Const PBM_SETBKCOLOR = CCM_SETBKCOLOR
Private Const WM_USER = &H400
Private Const PBM_SETBARCOLOR = (WM_USER + 9)
Private m_oCPULoad As CPULoad
Private m_lCPUs As Long
Private vbDGr As Long
Private vbLGr As Long
Private vbOWh As Long
Private OldTime As String
Private mBuffer As Long
Private mBufferDC As Long
Private mBlank As Long
Private mBlankDC As Long
Private BufferH As Long
Private BufferW As Long
Private Const PI As Single = 3.14159265358978
Sub ClearBuffer()
    BitBlt mBufferDC, 0, 0, BufferW, BufferH, mBlankDC, 0, 0, vbSrcCopy
End Sub

Sub BufferToScreen()
    BitBlt Me.hdc, 140, 110, BufferW, BufferH, mBufferDC, 0, 0, vbSrcCopy
End Sub

Sub CreateBlank()
    'create blank
    mBlankDC = CreateCompatibleDC(GetDC(0))
    mBlank = CreateCompatibleBitmap(GetDC(0), BufferW, BufferH)
    SelectObject mBlankDC, mBlank
    SetBkMode mBlankDC, 0
End Sub
Sub CreateBuffer()
    'create buffer (to go to destination hdc)
    mBufferDC = CreateCompatibleDC(GetDC(0))
    mBuffer = CreateCompatibleBitmap(GetDC(0), BufferW, BufferH)
    SelectObject mBufferDC, mBuffer
    SetBkMode mBufferDC, 0
End Sub

Private Sub Form_DblClick()
   Unload Me
   End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, _
    y As Single)
    Const WM_NCLBUTTONDOWN = &HA1
    Const HTCAPTION = 2
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub
Sub SetTopmostWindow(ByVal hWnd As Long, Optional topmost As Boolean = True)
    Const HWND_NOTOPMOST = -2
    Const HWND_TOPMOST = -1
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    SetWindowPos hWnd, IIf(topmost, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, _
        SWP_NOMOVE + SWP_NOSIZE
End Sub
Sub DrawProgressBar(frmIn As Form, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lBgCol As Long, lPbBg As Long, lPbFg As Long, lPerc As Long)
Dim hRPen As Long
Dim y As Long
Dim lPercEndPos As Long
Dim Point As PointAPI
    
    frmIn.ForeColor = lBgCol
    'draw background
    hRPen = CreateSolidBrush(lBgCol)
    SelectObject frmIn.hdc, hRPen
    Rectangle frmIn.hdc, X1, Y1, X2, Y2
    DeleteObject hRPen
    
    'blank out corner pixels
    SetPixel frmIn.hdc, X1, Y1, RGB(255, 0, 255)
    SetPixel frmIn.hdc, X2 - 1, Y1, RGB(255, 0, 255)
    SetPixel frmIn.hdc, X1, Y2 - 1, RGB(255, 0, 255)
    SetPixel frmIn.hdc, X2 - 1, Y2 - 1, RGB(255, 0, 255)
    
    If lPerc = 100 Then
        frmIn.ForeColor = &H6050FF
    Else
        frmIn.ForeColor = lPbBg
    End If
    For y = Y1 + 2 To Y2 - 3 Step 2
        Point.x = X1 + 2: Point.y = y
        MoveToEx frmIn.hdc, X1 + 2, y, Point
        LineTo frmIn.hdc, X2 - 2, y
    Next
    
    If lPerc = 0 Or lPerc = 100 Then GoTo LastBit
    
    lPercEndPos = (((X2 - X1 - 4) / 100) * lPerc) + X1 + 2
    
    If lPerc > 95 Then
        frmIn.ForeColor = &H6050FF
    ElseIf lPerc > 85 Then
        frmIn.ForeColor = &HC0FFFF
    Else
        frmIn.ForeColor = lPbFg
    End If
    
    For y = Y1 + 2 To Y2 - 3 Step 2
        Point.x = X1 + 2: Point.y = y
        MoveToEx frmIn.hdc, X1 + 2, y, Point
        LineTo frmIn.hdc, lPercEndPos, y
    Next
LastBit:
    Me.Refresh
    DoEvents
End Sub
Sub FrmTextOut(FormIn As Form, sIn As String, xPos As Integer, ypos As Integer, lColor As Long)
    SetTextColor FormIn.hdc, vbBlack
    TextOut FormIn.hdc, xPos, ypos + 1, sIn, Len(sIn)
    TextOut FormIn.hdc, xPos, ypos - 1, sIn, Len(sIn)
    TextOut FormIn.hdc, xPos - 1, ypos, sIn, Len(sIn)
    TextOut FormIn.hdc, xPos + 1, ypos, sIn, Len(sIn)
    
    SetTextColor FormIn.hdc, lColor
    TextOut FormIn.hdc, xPos, ypos, sIn, Len(sIn)

End Sub
Private Sub Form_Load()
Dim x As Long, y As Long
   
   App.Title = "Fosters Desktop Info"

    If Len(GetSetting(App.Title, "Settings", "Posx")) > 0 Then
        x = CLng(GetSetting(App.Title, "Settings", "Posx"))
    Else
        x = Screen.Width - (Me.Width * 1.2)
    End If
    If Len(GetSetting(App.Title, "Settings", "Posy")) > 0 Then
        y = CLng(GetSetting(App.Title, "Settings", "Posy"))
    Else
        y = (Me.Width * 0.5)
    End If
    If x < 0 Or x > Screen.Width Then x = Screen.Width - (Me.Width * 1.2)
    If y < 0 Or y > Screen.Height Then y = (Me.Width * 0.5)
    Me.Top = y
    Me.Left = x

   'for a dot shade background
   'For Y = 1 To Me.ScaleHeight Step 2
   '   For X = 1 To Me.ScaleWidth Step 2
   '      SetPixel Me.hdc, X, Y, 0
   '   Next
   'Next

   vbLGr = RGB(100, 200, 130)
   vbDGr = RGB(30, 60, 30)
   vbOWh = RGB(220, 220, 220)
   
   InitCPU
   
   FrmTextOut Me, "Physical RAM", 5, 5, vbOWh
   FrmTextOut Me, "Virtual RAM", 5, 25, vbOWh
   FrmTextOut Me, "CPU", 5, 45, vbOWh
   FrmTextOut Me, "Host IP", 5, 70, vbOWh
   FrmTextOut Me, "Host Name", 5, 85, vbOWh
   
   FrmTextOut Me, GetIPAddress, 100, 70, vbOWh
   FrmTextOut Me, GetIPHostName, 100, 85, vbOWh
   
   Timer1_Timer

   
   'SetTopmostWindow Me.hWnd

   BufferW = 60:    BufferH = BufferW
   
   CreateBlank
   DrawBlank
   CreateBuffer
   
   ShowTime
   
   DrawCalendar
   
   Me.Refresh
   SetColorTransparent Me, RGB(255, 0, 255)
   Timer1.Enabled = True
End Sub
Sub DrawCalendar()
Dim Stamp As New clsCalendarStamp
With Stamp
    .Background = vbBlack
    .BackgroundTrimIT = border
    
    .TrimITDepth = 1
    
    .Left = 15
    .Top = 180
    
    .CalendarMonth = Month(Now)
    .CalendarYear = Year(Now)
    
    .TargetImage = Me
    
    .DayBold = True
    .DayColor = RGB(230, 230, 230)
    .DayFont = "MS Sans Serif"
    .DayFontSize = 8
    
    .LabelBold = True
    .LabelColor = RGB(255, 255, 220)
    .LabelFont = "MS Sans Serif"
    .LabelFontSize = 8
    
    .TitleBold = True
    .TitleColor = RGB(230, 230, 255)
    .TitleFont = "MS Sans Serif"
    .TitleFontSize = 10
    
    .TodayColor = RGB(255, 130, 155)
    
    .DrawCalendar
End With

End Sub
Function GetMEMORY() As Long()
Dim memsts As MEMORYSTATUS
Dim RetMem(2) As Long


    GlobalMemoryStatus memsts
    RetMem(0) = Int((100 / memsts.dwTotalPhys) * (memsts.dwTotalPhys - memsts.dwAvailPhys))
    RetMem(1) = Int((100 / memsts.dwTotalVirtual) * (memsts.dwTotalVirtual - memsts.dwAvailVirtual))

    GetMEMORY = RetMem
End Function



Private Sub InitCPU()
Dim i As Long
Const lOffset As Long = 30
    
    Set m_oCPULoad = New CPULoad
    m_lCPUs = m_oCPULoad.GetCPUCount
        
End Sub

Private Function ReturnCPUPercent(lCPU As Long) As Single
    m_oCPULoad.CollectCPUData
    ReturnCPUPercent = m_oCPULoad.GetCPUUsage(lCPU)
End Function

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Settings", "Posx", Me.Left
    SaveSetting App.Title, "Settings", "Posy", Me.Top
    
    Set m_oCPULoad = Nothing
    DeleteObject mBlank: DeleteObject mBlankDC
    DeleteObject mBuffer: DeleteObject mBufferDC
End Sub



Private Sub Timer1_Timer()
Dim rMEM() As Long
Dim lL As Long
Dim lR As Long
   lL = 100: lR = 200
   
   rMEM = GetMEMORY
   DrawProgressBar Me, lL, 5, lR, 18, 0, vbDGr, vbLGr, rMEM(0)
   DrawProgressBar Me, lL, 25, lR, 38, 0, vbDGr, vbLGr, rMEM(1)
   DrawProgressBar Me, lL, 45, lR, 58, 0, vbDGr, vbLGr, ReturnCPUPercent(1)
   
   'FrmTextOut Me, Format(Now, "Mmmm dd YYYY"), 5, 110, vbOWh
   
   ShowTime
   
End Sub
Sub ShowTime()
   ClearBuffer
   ClockToBuffer
   BufferToScreen
End Sub
Sub DrawBlank()
Dim m_line As New LineGS
Dim m_Grad As New clsGradient
Dim BM As BITMAP

Dim mDl As Long
Dim An As Single
   With m_Grad
       .Angle = 130
       .Color2 = vbWhite 'RGB(150, 200, 255)
       .Color1 = RGB(180, 200, 255)
       .PictureHDC = mBlankDC
       .PictureHWND = GetObject(mBlank, Len(BM), BM)
       .Draw BufferW, BufferH
   End With
   
   mDl = (BufferW \ 2) - 1
   
   For An = 0 To 359 Step 30
       SetPixel mBlankDC, mDl + GimmeX(An, mDl * 0.8), mDl + GimmeY(An, mDl * 0.8), 0
   Next
   With m_line
      .LineGP mBlankDC, 0, BufferH - 1, BufferW, BufferH - 1, 0
      .LineGP mBlankDC, BufferW - 1, 0, BufferW - 1, BufferH, 0
   End With
End Sub
Sub ClockToBuffer()
Dim m_line As New LineGS
Dim hh As Long
Dim mm As Long
Dim ss As Long
Dim mDl As Long
Dim ssAng As Single
   hh = Format(Now, "hh"): If hh >= 12 Then hh = hh - 12
   mm = Format(Now, "nn")
   ss = Format(Now, "ss")
   ssAng = 180 - (CSng(ss * 6))
   mDl = (BufferW \ 2) - 1
   With m_line
       .LineGP mBufferDC, mDl, mDl, mDl + GimmeX(ssAng, mDl * 0.9), _
                                    mDl + GimmeY(ssAng, mDl * 0.9), RGB(100, 100, 100)
       .LineGP mBufferDC, mDl, mDl, mDl + GimmeX(180 - (CSng(((hh * 60) + mm) * 0.5)), mDl * 0.6), _
                                    mDl + GimmeY(180 - (CSng(((hh * 60) + mm) * 0.5)), mDl * 0.6), 0
       .LineGP mBufferDC, mDl, mDl, mDl + GimmeX(180 - CSng(mm * 6), mDl * 0.9), mDl + GimmeY(180 - CSng(mm * 6), mDl * 0.9), 0
       .CircleGP mBufferDC, mDl, mDl, 4, 4, 0
       .CircleGP mBufferDC, mDl, mDl, 2, 2, 0
   End With
End Sub
Function GimmeX(ByVal aIn As Single, lIn As Long) As Long
    GimmeX = sIn(aIn * (PI / 180)) * lIn
End Function
Function GimmeY(ByVal aIn As Single, lIn As Long) As Long
    GimmeY = Cos(aIn * (PI / 180)) * lIn
End Function


