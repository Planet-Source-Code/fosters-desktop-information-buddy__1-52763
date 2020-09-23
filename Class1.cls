VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCalendarStamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private m_xPos As Long
Private m_yPos As Long
Private m_BorderColor As Long
Private m_Month As Integer
Private m_Year As Integer
Private m_Destination As Form
Private m_TitleFont As String
Private m_TitleColor As Long
Private m_TitleFontSize As Long
Private m_TitleBold As Boolean
Private m_LabelFont As String
Private m_LabelColor As Long
Private m_LabelFontSize As Long
Private m_LabelBold As Boolean
Private m_DayFont As String
Private m_DayColor As Long
Private m_DayFontSize As Long
Private m_DayBold As Boolean
Private m_TodayColor As Long
Private m_TrimIT As Integer
Private m_TrimITDepth As Long
Public Enum TrimIT
    None = 0
    border = 1
    Dropshadow = 2
End Enum
'Let properties
Public Property Let BackgroundTrimIT(iT As TrimIT)
    m_TrimIT = iT
End Property
Public Property Let TrimITDepth(lDSD As Long)
    m_TrimITDepth = lDSD
End Property
Public Property Let Left(xPos As Long)
    m_xPos = xPos
End Property
Public Property Let Top(ypos As Long)
    m_yPos = ypos
End Property
Public Property Let Background(bgCol As Long)
    m_BorderColor = bgCol
End Property
Public Property Let CalendarMonth(iMonth As Integer)
    m_Month = iMonth
End Property
Public Property Let CalendarYear(iYear As Integer)
    m_Year = iYear
End Property
Friend Property Let TargetImage(ByRef ctlDestination As Form)
    Set m_Destination = ctlDestination
End Property
Public Property Let TitleFont(lTF As String)
    m_TitleFont = lTF
End Property
Public Property Let TitleColor(lTC As Long)
    m_TitleColor = lTC
End Property
Public Property Let TitleFontSize(lTFS As Long)
    m_TitleFontSize = lTFS
End Property
Public Property Let TitleBold(lTB As Long)
    m_TitleBold = lTB
End Property
Public Property Let LabelFont(lLF As String)
    m_LabelFont = lLF
End Property
Public Property Let LabelColor(lLC As Long)
    m_LabelColor = lLC
End Property
Public Property Let LabelFontSize(lLFS As Long)
    m_LabelFontSize = lLFS
End Property
Public Property Let LabelBold(lLB As Long)
    m_LabelBold = lLB
End Property
Public Property Let DayFont(lDF As String)
    m_DayFont = lDF
End Property
Public Property Let DayColor(lDC As Long)
    m_DayColor = lDC
End Property
Public Property Let DayFontSize(lDFS As Long)
    m_DayFontSize = lDFS
End Property
Public Property Let DayBold(lDB As Long)
    m_DayBold = lDB
End Property
Public Property Let TodayColor(lTC As Long)
    m_TodayColor = lTC
End Property

Public Sub DrawCalendar()
    'm_bordercolor Outline
    If m_TrimIT = 1 Then
        PrintCalendar m_Destination, m_Month, m_Year, m_xPos - m_TrimITDepth, m_yPos, m_TitleFont, m_BorderColor, m_TitleFontSize, m_TitleBold, m_LabelFont, m_BorderColor, m_LabelFontSize, m_LabelBold, m_DayFont, m_BorderColor, m_DayFontSize, m_DayBold, m_BorderColor
        PrintCalendar m_Destination, m_Month, m_Year, m_xPos + m_TrimITDepth, m_yPos, m_TitleFont, m_BorderColor, m_TitleFontSize, m_TitleBold, m_LabelFont, m_BorderColor, m_LabelFontSize, m_LabelBold, m_DayFont, m_BorderColor, m_DayFontSize, m_DayBold, m_BorderColor
        PrintCalendar m_Destination, m_Month, m_Year, m_xPos, m_yPos - m_TrimITDepth, m_TitleFont, m_BorderColor, m_TitleFontSize, m_TitleBold, m_LabelFont, m_BorderColor, m_LabelFontSize, m_LabelBold, m_DayFont, m_BorderColor, m_DayFontSize, m_DayBold, m_BorderColor
        PrintCalendar m_Destination, m_Month, m_Year, m_xPos, m_yPos + m_TrimITDepth, m_TitleFont, m_BorderColor, m_TitleFontSize, m_TitleBold, m_LabelFont, m_BorderColor, m_LabelFontSize, m_LabelBold, m_DayFont, m_BorderColor, m_DayFontSize, m_DayBold, m_BorderColor
        PrintCalendar m_Destination, m_Month, m_Year, m_xPos - m_TrimITDepth, m_yPos + m_TrimITDepth, m_TitleFont, m_BorderColor, m_TitleFontSize, m_TitleBold, m_LabelFont, m_BorderColor, m_LabelFontSize, m_LabelBold, m_DayFont, m_BorderColor, m_DayFontSize, m_DayBold, m_BorderColor
        PrintCalendar m_Destination, m_Month, m_Year, m_xPos + m_TrimITDepth, m_yPos - m_TrimITDepth, m_TitleFont, m_BorderColor, m_TitleFontSize, m_TitleBold, m_LabelFont, m_BorderColor, m_LabelFontSize, m_LabelBold, m_DayFont, m_BorderColor, m_DayFontSize, m_DayBold, m_BorderColor
        PrintCalendar m_Destination, m_Month, m_Year, m_xPos - m_TrimITDepth, m_yPos - m_TrimITDepth, m_TitleFont, m_BorderColor, m_TitleFontSize, m_TitleBold, m_LabelFont, m_BorderColor, m_LabelFontSize, m_LabelBold, m_DayFont, m_BorderColor, m_DayFontSize, m_DayBold, m_BorderColor
        PrintCalendar m_Destination, m_Month, m_Year, m_xPos + m_TrimITDepth, m_yPos + m_TrimITDepth, m_TitleFont, m_BorderColor, m_TitleFontSize, m_TitleBold, m_LabelFont, m_BorderColor, m_LabelFontSize, m_LabelBold, m_DayFont, m_BorderColor, m_DayFontSize, m_DayBold, m_BorderColor
    End If
    
    'dropshadow in m_bordercolor
    If m_TrimIT = 2 Then
        PrintCalendar m_Destination, _
                      m_Month, m_Year, _
                      m_xPos + m_TrimITDepth, m_yPos + m_TrimITDepth, _
                      m_TitleFont, m_BorderColor, m_TitleFontSize, m_TitleBold, _
                      m_LabelFont, m_BorderColor, m_LabelFontSize, m_LabelBold, _
                      m_DayFont, m_BorderColor, m_DayFontSize, m_DayBold, _
                      m_BorderColor
    End If
    
    'Text
    PrintCalendar m_Destination, _
                 m_Month, m_Year, _
                 m_xPos, m_yPos, _
                 m_TitleFont, m_TitleColor, m_TitleFontSize, m_TitleBold, _
                 m_LabelFont, m_LabelColor, m_LabelFontSize, m_LabelBold, _
                 m_DayFont, m_DayColor, m_DayFontSize, m_DayBold, _
                 m_TodayColor

End Sub
Private Sub PrintCalendar(ByRef picIn As Form, iMonth As Integer, iYear As Integer, _
                          xPos As Long, ypos As Long, _
                          TitleFont As String, TitleColor As Long, TitleFontSize As Long, TitleBold As Boolean, _
                          LabelFont As String, LabelColor As Long, LabelFontSize As Long, LabelBold As Boolean, _
                          DayFont As String, DayColor As Long, DayFontSize As Long, DayBold As Boolean, _
                          TodayColor As Long)
Dim sText As String
Dim x As Long
Dim LabelColWidth As Long
Dim LabelTop As Long
Dim DayTop As Long
Dim DayHeight As Long
Dim DayPosX As Long
Dim DayPosY As Long
Dim CurrDayNum As Long
Dim xOffset As Long
    'deduce colwidth
    picIn.Font = LabelFont
    picIn.FontSize = LabelFontSize
    picIn.FontBold = LabelBold
    LabelColWidth = picIn.TextWidth("WW")
    picIn.Font = DayFont
    picIn.FontSize = DayFontSize
    picIn.FontBold = DayBold
    If picIn.TextWidth("WW") > LabelColWidth Then
        LabelColWidth = picIn.TextWidth("WW")
    End If
    LabelColWidth = LabelColWidth + 1

    'month label
    picIn.Font = TitleFont
    picIn.FontSize = TitleFontSize
    
    sText = Format(DateSerial(iYear, iMonth, 1), "Mmmm  yyyy")
    picIn.FontBold = TitleBold
    
    SetTextColor picIn.hdc, TitleColor
    If picIn.TextWidth(sText) > (LabelColWidth * 7) Then
        xOffset = picIn.TextWidth(sText) - (LabelColWidth * 7)
        TextOut picIn.hdc, xPos, ypos, sText, Len(sText)
    Else
        xOffset = 0
        TextOut picIn.hdc, xPos + 2 + (LabelColWidth * 7) - picIn.TextWidth(sText), ypos, sText, Len(sText)
    End If
    
    
    LabelTop = picIn.TextHeight(sText) + ypos + 6
    
    
    'day labels
    picIn.Font = LabelFont
    picIn.FontSize = LabelFontSize
    picIn.FontBold = LabelBold
    
    For x = 1 To 7
        sText = Format(x, "Ddd")
        SetTextColor picIn.hdc, LabelColor
        TextOut picIn.hdc, xOffset + (xPos + ((x - 1) * LabelColWidth) + 8) + LabelColWidth - picIn.TextWidth(sText) - 3, LabelTop, Mid(sText, 1, 2), 2
    Next x
    
    DayTop = picIn.TextHeight(sText) + LabelTop + 2
    
    'day numerics
    picIn.Font = DayFont
    picIn.FontSize = DayFontSize
    picIn.FontBold = DayBold
    
    DayHeight = picIn.TextHeight("999")
    
    DayPosX = xPos + (Offset(iMonth, iYear) * LabelColWidth)
    DayPosY = DayTop
    CurrDayNum = 1
    Do
        sText = CurrDayNum
        SetTextColor picIn.hdc, IIf(iMonth = Month(Now) And iYear = Year(Now) And CurrDayNum = Day(Now), TodayColor, DayColor)
        TextOut picIn.hdc, xOffset + DayPosX + LabelColWidth - picIn.TextWidth(sText), DayPosY, sText, Len(sText)
        CurrDayNum = CurrDayNum + 1
        If Format(DateSerial(iYear, iMonth, CurrDayNum), "Ddd") = "Sun" Then
            DayPosX = xPos
            DayPosY = DayPosY + DayHeight
        Else
            DayPosX = DayPosX + LabelColWidth
        End If
    Loop Until CurrDayNum > DaysInMonth(iMonth, iYear)
    
End Sub
Private Function DaysInMonth(iM As Integer, iY As Integer) As Integer
Dim dteStart As Date
Dim dteEnd As Date
    dteStart = DateSerial(iY, iM, 1)
    dteEnd = DateAdd("m", 1, dteStart)
    DaysInMonth = DateDiff("d", dteStart, dteEnd)
End Function
Private Function Offset(iM As Integer, iY As Integer) As Integer
Dim sDte As String
    Offset = 0
    sDte = DateSerial(iY, iM, 1)
    Select Case Format(sDte, "Ddd")
        Case "Sun"
        Offset = 0
        Case "Mon"
        Offset = 1
        Case "Tue"
        Offset = 2
        Case "Wed"
        Offset = 3
        Case "Thu"
        Offset = 4
        Case "Fri"
        Offset = 5
        Case "Sat"
        Offset = 6
    End Select
End Function

