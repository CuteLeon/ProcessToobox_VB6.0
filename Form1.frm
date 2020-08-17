VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "小眼进程工具箱"
   ClientHeight    =   3570
   ClientLeft      =   12600
   ClientTop       =   4710
   ClientWidth     =   4500
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   238
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      Caption         =   "关闭进程"
      Height          =   315
      Left            =   120
      TabIndex        =   19
      Top             =   3180
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
      Caption         =   "文件属性"
      Height          =   315
      Left            =   1620
      TabIndex        =   18
      Top             =   3180
      Width           =   1155
   End
   Begin VB.TextBox PIDText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   2400
      TabIndex        =   16
      Top             =   2280
      Width           =   1995
   End
   Begin VB.CommandButton Command1 
      Caption         =   "定位到文件"
      Height          =   315
      Left            =   2880
      TabIndex        =   15
      Top             =   3180
      Width           =   1515
   End
   Begin VB.TextBox EXEPath 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   120
      TabIndex        =   13
      Top             =   2820
      Width           =   4275
   End
   Begin VB.TextBox FatherText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   2400
      TabIndex        =   11
      Top             =   840
      Width           =   1995
   End
   Begin VB.TextBox FatherhWnd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   2400
      TabIndex        =   10
      Top             =   480
      Width           =   1995
   End
   Begin VB.TextBox PasswordText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   2400
      TabIndex        =   4
      Top             =   1920
      Width           =   1995
   End
   Begin VB.TextBox WndClassText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   1995
   End
   Begin VB.TextBox PointText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   1995
   End
   Begin VB.TextBox hWndText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   1995
   End
   Begin VB.Frame Frame2 
      Caption         =   " 拖动图标"
      Height          =   1395
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
      Begin VB.Image Image3 
         Height          =   720
         Left            =   240
         Picture         =   "Form1.frx":57E2
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "PID："
      Height          =   180
      Left            =   1920
      TabIndex        =   17
      Top             =   2340
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "程序路径："
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   2580
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "顶级标题："
      Height          =   180
      Left            =   1500
      TabIndex        =   12
      Top             =   900
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "顶级句柄："
      Height          =   180
      Left            =   1500
      TabIndex        =   9
      Top             =   540
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "密码文本："
      Height          =   180
      Left            =   1500
      TabIndex        =   8
      Top             =   1980
      Width           =   900
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "目标类型："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1500
      TabIndex        =   7
      Top             =   1620
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "鼠标坐标："
      Height          =   180
      Left            =   1500
      TabIndex        =   6
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "句柄："
      Height          =   180
      Left            =   1860
      TabIndex        =   5
      Top             =   180
      Width           =   540
   End
   Begin VB.Menu TextMenu 
      Caption         =   "TextMenu"
      Visible         =   0   'False
      Begin VB.Menu LockIt 
         Caption         =   "锁定"
      End
      Begin VB.Menu HideIt 
         Caption         =   "隐藏"
      End
      Begin VB.Menu MoveIt 
         Caption         =   "调整"
      End
      Begin VB.Menu TopNot 
         Caption         =   "置前/置后"
      End
      Begin VB.Menu Attributes 
         Caption         =   "(半)透明"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()   '注册并初始化通用控件窗口类
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Const WM_GETTEXT = &HD
Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400

Private Type SHELLEXECUTEINFO
       cbSize As Long
       fMask As Long
       hwnd As Long
       lpVerb As String
       lpFile As String
       lpParameters As String
       lpDirectory As String
       nShow As Long
       hInstApp As Long
       lpIDList As Long
       lpClass As String
       hkeyClass As Long
       dwHotKey As Long
       hIcon As Long
       hProcess As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Dim HandIng As Boolean

Private Sub Command1_Click()
  If Dir(Trim(EXEPath)) <> "" Then Shell "explorer.exe /select," & EXEPath.Text, vbNormalFocus
End Sub

Private Sub Command2_Click()
  If Dir(Trim(EXEPath)) <> "" Then ShowProperties Trim(EXEPath), Me.hwnd
End Sub

Private Sub Command3_Click()
  On Error Resume Next
  If PIDText = "" Then Exit Sub
  Dim LnghWndProcess As Long, Hand As Long, ExitCode As Long
  Hand = OpenProcess(&H1, True, Val(PIDText))
  TerminateProcess Hand, ExitCode
  CloseHandle Hand
End Sub

Private Sub hWndText_Change()
If HandIng = True Or hWndText = "" Then Exit Sub
  Dim Rtn As Long
  Dim TempStr As String
  Dim StrLong As Long
  Dim TemphWnd As Long, LasthWnd As Long
  TemphWnd = CLng(hWndText)
  Do Until TemphWnd = 0
    LasthWnd = TemphWnd
    TemphWnd = GetParent(LasthWnd)
  Loop
  FatherhWnd.Text = LasthWnd
  
  Dim FatherPID As Long
  Rtn = GetWindowThreadProcessId(LasthWnd, FatherPID)
  PIDText.Text = FatherPID
  
  EXEPath.Text = GetProcessPathByProcessID(PIDText)
    
  TempStr = Space(255)
  StrLong = Len(TempStr)
  Rtn = SendMessage(LasthWnd, WM_GETTEXT, StrLong, TempStr)
  TempStr = Trim(TempStr)
  FatherText.Text = TempStr
  
  TempStr = Space(255)
  StrLong = Len(TempStr)
  Rtn = GetClassName(hWndText, TempStr, StrLong)
  If Rtn = 0 Then Exit Sub
  TempStr = Trim(TempStr)
  WndClassText.Text = TempStr
  
  TempStr = Space(255)
  StrLong = Len(TempStr)
  Rtn = SendMessage(hWndText, WM_GETTEXT, StrLong, TempStr)
  TempStr = Trim(TempStr)
  PasswordText.Text = TempStr
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  HandIng = True
  Screen.MouseIcon = Image3.Picture
  Screen.MousePointer = vbCustom
  SetCapture (Me.hwnd)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Screen.MousePointer = vbDefault
  ReleaseCapture
  HandIng = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If HandIng = True Then
    Dim Rtn As Long, CurWnd As Long
    Dim TempStr As String
    Dim StrLong As Long
    Dim Cpos As String
    Dim Point As POINTAPI
    Point.X = X
    Point.Y = Y

    If ClientToScreen(Me.hwnd, Point) = 0 Then Exit Sub
    CurWnd = WindowFromPoint(Point.X, Point.Y)
    hWndText.Text = Trim(Str(CurWnd))
    
    Dim TemphWnd As Long, LasthWnd As Long
    TemphWnd = CLng(hWndText)
    Do Until TemphWnd = 0
      LasthWnd = TemphWnd
      TemphWnd = GetParent(LasthWnd)
    Loop
    FatherhWnd.Text = LasthWnd
    
    Dim FatherPID As Long
    Rtn = GetWindowThreadProcessId(LasthWnd, FatherPID)
    PIDText.Text = FatherPID
    
    EXEPath.Text = GetProcessPathByProcessID(PIDText)
    
    TempStr = Space(255)
    StrLong = Len(TempStr)
    Rtn = SendMessage(LasthWnd, WM_GETTEXT, StrLong, TempStr)
    TempStr = Trim(TempStr)
    FatherText.Text = TempStr
    Cpos = Trim(Str(Point.X)) & "," & Trim(Str(Point.Y))
    PointText.Text = Cpos
    
    TempStr = Space(255)
    StrLong = Len(TempStr)
    Rtn = GetClassName(CurWnd, TempStr, StrLong)
    If Rtn = 0 Then Exit Sub
    TempStr = Trim(TempStr)
    WndClassText.Text = TempStr
    
    TempStr = Space(255)
    StrLong = Len(TempStr)
    Rtn = SendMessage(CurWnd, WM_GETTEXT, StrLong, TempStr)
    TempStr = Trim(TempStr)
    PasswordText.Text = TempStr
  End If
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Function GetProcessPathByProcessID(pid As Long) As String               '获取应用程序的完整路径
  On Error Resume Next
  Dim cbNeeded As Long
  Dim szBuf(1 To 250) As Long
  Dim Ret As Long
  Dim szPathName As String
  Dim nSize As Long
  Dim hProcess As Long
  
  hProcess = OpenProcess(&H400 Or &H10, 0, pid)
  If hProcess <> 0 Then
    Ret = EnumProcessModules(hProcess, szBuf(1), 250, cbNeeded)
    If Ret <> 0 Then
      szPathName = Space$(260)
      nSize = 500
      Ret = GetModuleFileNameExA(hProcess, szBuf(1), szPathName, nSize)
      GetProcessPathByProcessID = Left$(szPathName, Ret)
    End If
  End If
  Ret = CloseHandle(hProcess)
End Function

Public Function ShowProperties(filename As String, OwnerhWnd As Long) As Long
  Dim SEI As SHELLEXECUTEINFO
  Dim Ret As Long
  With SEI
    .cbSize = Len(SEI)
    .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
    .hwnd = OwnerhWnd
    .lpVerb = "properties"
    .lpFile = filename
    .lpParameters = vbNullChar
    .lpDirectory = vbNullChar
    .nShow = 0
    .hInstApp = 0
    .lpIDList = 0
  End With
  Ret = ShellExecuteEX(SEI)
  ShowProperties = SEI.hInstApp
End Function

Private Sub PIDText_Change()
If HandIng = True Or PIDText = "" Then Exit Sub
  EXEPath.Text = GetProcessPathByProcessID(PIDText)
End Sub
