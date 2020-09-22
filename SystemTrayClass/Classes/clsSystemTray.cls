VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystemTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'SystemTray Class
'
'Author Ben Vonk
'23-09-2010 First version
'27-09-2010 Second version Add Balloon events
'02-10-2010 Third version Add Balloon timer and fixed some bugs
'06-11-2010 Fourth version fixed some bugs
'09-11-2010 Fifth version Add hWnd function and make some changes
'06-12-2010 Sixth version Add ReceivedData event

Option Explicit

' Public Events
Public Event BalloonClick()
Public Event BalloonClose()
Public Event BalloonHide()
Public Event BalloonShow()
Public Event BalloonTimeOut()
Public Event Click(Button As Integer)
Public Event DblClick(Button As Integer)
Public Event MouseDown(Button As Integer)
Public Event MouseMove()
Public Event MouseUp(Button As Integer)
Public Event ReceivedData(Data As String)

' Private Constants
Private Const NOTIFYICONDATA_V1_SIZE As Long = 88
Private Const CLASS_NAME             As String = "SystemTray_Class"

' Public Enumeration
Public Enum BalloonIconConstants
   NIIF_NONE = &H0
   NIIF_INFO = &H1
   NIIF_WARNING = &H2
   NIIF_ERROR = &H3
   NIIF_ICON = &H4
   NIIF_GUID = &H5
   NIIF_ASK = &H7
   NIIF_ICON_MASK = &HF
End Enum

' Private Types
Private Type GUID
   Data1                             As Long
   Data2                             As Integer
   Data3                             As Integer
   Data4(7)                          As Byte
End Type

Private Type NotifyIconData
   cbSize                            As Long
   hWnd                              As Long
   uID                               As Long
   uFlags                            As Long
   uCallbackMessage                  As Long
   hIcon                             As Long
   szTip                             As String * 128
   dwState                           As Long
   dwStateMask                       As Long
   szInfo                            As String * 256
   uTimeoutAndVersion                As Long
   szInfoTitle                       As String * 64
   dwInfoFlags                       As Long
   guidItem                          As GUID
End Type

Private Type PointAPI
   X                                 As Long
   Y                                 As Long
End Type

Private Type Rect
   Left                              As Long
   Top                               As Long
   Right                             As Long
   Bottom                            As Long
End Type

Private Type WindowClass
   style                             As Long
   lpfnwndproc                       As Long
   cbClsextra                        As Long
   cbWndExtra2                       As Long
   hInstance                         As Long
   hIcon                             As Long
   hCursor                           As Long
   hbrBackground                     As Long
   lpszMenuName                      As Long
   lpszClassName                     As String
End Type

' Private Variables
Private m_BalloonClicked             As Boolean
Private m_BalloonClosed              As Boolean
Private m_BalloonShowed              As Boolean
Private m_Enabled                    As Boolean
Private m_Visible                    As Boolean
Private m_hWndHidden                 As Long
Private m_hWndMenu                   As Long
Private m_hWndParent                 As Long
Private m_SizeNID                    As Long
Private m_TimerID                    As Long
Private m_NID                        As NotifyIconData

' Private API's
Private Declare Function ShellNotifyIcon Lib "Shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NotifyIconData) As Long
Private Declare Function CreateWindowEx Lib "User32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function GetCursorPos Lib "User32" (lpPoint As PointAPI) As Long
Private Declare Function GetSystemMenu Lib "User32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function KillTimer Lib "User32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RegisterClass Lib "User32" Alias "RegisterClassA" (Class As WindowClass) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetForegroundWindow Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function SetTimer Lib "User32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function TrackPopupMenu Lib "User32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As Rect) As Long
Private Declare Function UnregisterClass Lib "User32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "Version" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerQueryValue Lib "Version" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lpBuffer As Any, nVerSize As Long) As Long

Public Property Get Enabled() As Boolean

   Enabled = m_Enabled

End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)

   m_Enabled = NewEnabled

End Property

Public Property Get Icon() As Long

   Icon = m_NID.hIcon

End Property

Public Property Let Icon(ByVal NewIcon As Long)

   m_NID.hIcon = NewIcon
   
   Call ModifyIcon(m_BalloonShowed)

End Property

Public Property Get Menu() As Long

   Menu = m_hWndMenu

End Property

Public Property Let Menu(ByVal NewMenu As Long)

   m_hWndMenu = NewMenu

End Property

Public Property Get Parent() As Long

   Parent = m_hWndParent

End Property

Public Property Let Parent(ByRef NewParent As Long)

   m_hWndParent = NewParent

End Property

Public Property Get TipText() As String

   TipText = m_NID.szTip

End Property

Public Property Let TipText(ByVal NewTipText As String)

   m_NID.szTip = NewTipText & vbNullChar
   
   Call ModifyIcon(m_BalloonShowed)

End Property

Public Function hWnd() As Long

   hWnd = m_hWndHidden

End Function

Public Function Visible() As Boolean

   Visible = m_Visible

End Function

Public Sub AddIcon()

Const NIM_ADD As Long = &H0

   Debug.Assert m_Visible = False
   ShellNotifyIcon NIM_ADD, m_NID
   m_Visible = True

End Sub

Public Sub DeleteIcon()

Const NIM_DELETE As Long = &H2

   If m_TimerID Then m_TimerID = KillTimer(m_hWndParent, ObjPtr(Me) + 1)
   
   Debug.Assert m_Visible
   ShellNotifyIcon NIM_DELETE, m_NID
   m_Visible = False

End Sub

Public Sub HideBalloon()

   If m_TimerID > 0 Then If KillTimer(m_hWndHidden, ObjPtr(Me) + 1) = 1 Then m_TimerID = 0
   If Not m_BalloonShowed Then Exit Sub
   
   With m_NID
      .dwInfoFlags = 0
      .szInfoTitle = vbNullChar
      .szInfo = vbNullChar
      m_BalloonShowed = False
   End With
   
   Call ModifyIcon

End Sub

Public Sub RecreateIcon()

   If m_Visible Then
      m_Visible = False
      
      Call AddIcon
   End If

End Sub

Public Sub ShowBalloon(ByVal Title As String, ByVal Info As String, Optional ByVal BalloonIcon As BalloonIconConstants = NIIF_INFO, Optional ByVal TimeOut As Long, Optional ByVal Sound As Boolean)

Const NIF_INFO     As Long = &H10
Const NIIF_NOSOUND As Long = &H10

   If Not m_Enabled Or (m_SizeNID = NOTIFYICONDATA_V1_SIZE) Then Exit Sub
   
   With m_NID
      .uFlags = .uFlags Or NIF_INFO
      .cbSize = m_SizeNID
      .dwInfoFlags = BalloonIcon Or (NIIF_NOSOUND And Not Sound)
      .szInfoTitle = Title & vbNullChar
      .szInfo = Info & vbNullChar
      m_BalloonClicked = False
      m_BalloonClosed = False
      m_BalloonShowed = True
   End With
   
   Call ModifyIcon
   
   If (TimeOut > 0) And (m_TimerID = 0) Then m_TimerID = SetTimer(m_hWndHidden, ObjPtr(Me) + 1, TimeOut, AddressOf SystemTrayWndProc)

End Sub

Friend Sub ProcessMessage(ByVal ID As Long, ByVal MouseEvent As Long, Optional ByVal SendData As String)

Const NIN_BALLOONHIDE      As Long = &H403
Const NIN_BALLOONSHOW      As Long = &H402
Const NIN_BALLOONTIMEOUT   As Long = &H404
Const NIN_BALLOONUSERCLICK As Long = &H405
Const WM_LBUTTONDBLCLK     As Long = &H203
Const WM_LBUTTONDOWN       As Long = &H201
Const WM_LBUTTONUP         As Long = &H202
Const WM_MBUTTONDBLCLK     As Long = &H209
Const WM_MBUTTONDOWN       As Long = &H207
Const WM_MBUTTONUP         As Long = &H208
Const WM_MOUSEMOVE         As Long = &H200
Const WM_RBUTTONDBLCLK     As Long = &H206
Const WM_RBUTTONDOWN       As Long = &H204
Const WM_RBUTTONUP         As Long = &H205

   If m_Enabled Then
      If (m_TimerID > 0) And (m_TimerID = ID) Then
         If Not m_BalloonClicked And Not m_BalloonClosed Then
            Call HideBalloon
            
            RaiseEvent BalloonTimeOut
         End If
         
      Else
         Select Case MouseEvent
            Case NIN_BALLOONHIDE
               RaiseEvent BalloonHide
               
            Case NIN_BALLOONSHOW
               RaiseEvent BalloonShow
               
            Case NIN_BALLOONTIMEOUT
               Call HideBalloon
               
               m_BalloonClosed = True
               RaiseEvent BalloonClose
               
            Case NIN_BALLOONUSERCLICK
               Call HideBalloon
               
               m_BalloonClicked = True
               RaiseEvent BalloonClick
               
            Case WM_ACTIVATE
               If ID = 0 Then RaiseEvent ReceivedData(SendData)
               
            Case WM_LBUTTONDBLCLK
               RaiseEvent DblClick(vbLeftButton)
               
            Case WM_LBUTTONDOWN
               RaiseEvent MouseDown(vbLeftButton)
               
            Case WM_LBUTTONUP
               RaiseEvent MouseUp(vbLeftButton)
               RaiseEvent Click(vbLeftButton)
               
            Case WM_MBUTTONDBLCLK
               RaiseEvent DblClick(vbMiddleButton)
               
            Case WM_MBUTTONDOWN
               RaiseEvent MouseDown(vbMiddleButton)
               
            Case WM_MBUTTONUP
               RaiseEvent MouseUp(vbMiddleButton)
               RaiseEvent Click(vbMiddleButton)
               
            Case WM_MOUSEMOVE
               RaiseEvent MouseMove
               
            Case WM_RBUTTONDBLCLK
               RaiseEvent DblClick(vbRightButton)
               
            Case WM_RBUTTONDOWN
               RaiseEvent MouseDown(vbRightButton)
               
            Case WM_RBUTTONUP
               If m_hWndMenu Then
                  Call ShowMenu
                  
               Else
                  RaiseEvent MouseUp(vbRightButton)
                  RaiseEvent Click(vbRightButton)
               End If
         End Select
      End If
      
   Else
      SetForegroundWindow m_hWndParent
   End If

End Sub

Private Sub ModifyIcon(Optional ByVal Cancel As Boolean)

Const NIM_MODIFY As Long = &H1

   If Not m_Enabled Or Not m_Visible Or Cancel Then Exit Sub
   
   Debug.Assert m_Visible
   ShellNotifyIcon NIM_MODIFY, m_NID

End Sub

Private Sub ShowMenu()

Const TPM_RETURNCMD   As Long = &H100
Const TPM_RIGHTBUTTON As Long = &H2
Const TPM_TOPALIGN    As Long = &H0
Const WM_SYSCOMMAND   As Long = &H112&

Dim lngMenu           As Long
Dim lngMenuID         As Long
Dim ptaCursorPos      As PointAPI
Dim rctMenu           As Rect

   lngMenu = GetSystemMenu(m_hWndMenu, &H0&)
   
   If lngMenu = 0 Then Exit Sub
   
   GetCursorPos ptaCursorPos
   SetForegroundWindow m_hWndParent
   lngMenuID = TrackPopupMenu(lngMenu, TPM_RETURNCMD Or TPM_RIGHTBUTTON Or TPM_TOPALIGN, ptaCursorPos.X, ptaCursorPos.Y, 0, m_hWndMenu, rctMenu)
   
   If lngMenuID Then PostMessage m_hWndMenu, WM_SYSCOMMAND, lngMenuID, lngMenu

End Sub

Private Sub Class_Initialize()

Const NIF_ICON               As Long = &H2
Const NIF_MESSAGE            As Long = &H1
Const NIF_TIP                As Long = &H4
Const NOTIFYICONDATA_V2_SIZE As Long = 488
Const NOTIFYICONDATA_V3_SIZE As Long = 504
Const SHELL_32               As String = "Shell32"

Dim lngReturn                As Long
Dim lngVersion               As Long
Dim wndClass                 As WindowClass

   lngReturn = GetFileVersionInfoSize(SHELL_32, 0)
   
   If lngReturn > 0 Then
      ReDim bytBuffer(lngReturn - 1) As Byte
      
      GetFileVersionInfo SHELL_32, 0, lngReturn, bytBuffer(0)
      
      If VerQueryValue(bytBuffer(0), "\", lngReturn, 0) = 1 Then lngVersion = GetVersion(lngReturn + 10)
      
      Erase bytBuffer
   End If
   
   If lngVersion < 5 Then
      m_SizeNID = NOTIFYICONDATA_V1_SIZE
      
   ElseIf lngVersion = 6 Then
      m_SizeNID = NOTIFYICONDATA_V2_SIZE
      
   Else
      m_SizeNID = NOTIFYICONDATA_V3_SIZE
   End If
   
   With wndClass
      .lpfnwndproc = Pass(AddressOf SystemTrayWndProc)
      .hInstance = App.hInstance
      .lpszClassName = CLASS_NAME
   End With
   
   m_Enabled = True
   RegisterClass wndClass
   m_hWndHidden = CreateWindowEx(0, CLASS_NAME, CLASS_NAME_HIDDEN & App.Title, 0, 0, 0, 100, 100, 0, 0, 0, ByVal 0)
   SetWindowLong m_hWndHidden, GWL_USERDATA, CreateRef(Me)
   
   Call InitMessage
   
   With m_NID
      .cbSize = m_SizeNID
      .hWnd = m_hWndHidden
      .szTip = vbNullString
      .uCallbackMessage = WM_USER_SYSTRAY
      .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   End With

End Sub

Private Sub Class_Terminate()

Const WM_CLOSE As Long = &H10

   SendMessage m_hWndHidden, WM_CLOSE, 0, ByVal 0&
   
   If m_Visible Then DeleteIcon
   
   UnregisterClass CLASS_NAME, App.hInstance

End Sub

