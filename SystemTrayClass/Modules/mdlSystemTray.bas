Attribute VB_Name = "mdlSystemTray"
'SystemTray Module (Belonged to the SystemTray Class)
'
'Author Ben Vonk
'23-09-2010 First version
'27-09-2010 Second version Add Balloon events
'02-10-2010 Third version Add Balloon timer and fixed some bugs
'06-11-2010 Fourth version fixed some bugs
'09-11-2010 Fifth version Add hWnd function and make some changes
'06-12-2010 Sixth version Add ReceivedData event

Option Explicit

' Public Constants
Public Const GWL_USERDATA      As Long = -21
Public Const WM_ACTIVATE       As Long = &H6
Public Const WM_COPYDATA       As Long = &H4A
Public Const WM_USER_SYSTRAY   As Long = &H405
Public Const CLASS_NAME_HIDDEN As String = "SystemTray_HiddenWindow_"
Public Const RECEIVED_DATA     As String = "Received_Data_"

' Public Type
Public Type CopyDataStruct
   dwData                      As Long
   cbData                      As Long
   lpData                      As Long
End Type

' Private Variable
Private m_TaskbarRestart       As Long

' Private API's
Private Declare Function StrLen Lib "Kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function DefWindowProc Lib "User32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function RegisterWindowMessage Lib "User32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function CreateRef(ByRef cObject As clsSystemTray) As Long

   Call CopyMemory(ByVal VarPtr(CreateRef), ByVal VarPtr(cObject), 4)

End Function

Public Function GetVersion(ByVal nValue As Long) As Long

   Call CopyMemory(GetVersion, ByVal nValue, 2)

End Function

Public Function Pass(ByVal nValue As Long) As Long

   Pass = nValue

End Function

Public Function SystemTrayWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Const WM_TIMER          As Long = &H113

Dim bytBuffer(1 To 255) As Byte
Dim cstObject           As clsSystemTray
Dim cdsData             As CopyDataStruct
Dim strData             As String

   If (uMsg = WM_COPYDATA) And (wParam = WM_ACTIVATE) Then
      CopyMemory cdsData, ByVal lParam, Len(cdsData)
      
      With cdsData
         If .dwData = 3 Then
            CopyMemory bytBuffer(1), ByVal .lpData, .cbData
            strData = StrConv(bytBuffer, vbUnicode)
            
            If Left(strData, Len(RECEIVED_DATA)) = RECEIVED_DATA Then
               strData = Mid(strData, Len(RECEIVED_DATA) + 1)
               strData = Left(strData, StrLen(StrPtr(strData)))
               Set cstObject = DeRef(GetWindowLong(hWnd, GWL_USERDATA))
               
               Call cstObject.ProcessMessage(0, wParam, strData)
               
               DestroyRef VarPtr(cstObject)
            End If
         End If
      End With
      
   ElseIf (uMsg = WM_USER_SYSTRAY) Or (uMsg = WM_TIMER) Or (uMsg = m_TaskbarRestart) Then
      Set cstObject = DeRef(GetWindowLong(hWnd, GWL_USERDATA))
      
      If uMsg = m_TaskbarRestart Then
         Call cstObject.RecreateIcon
         
      Else
         Call cstObject.ProcessMessage(wParam, lParam)
      End If
      
      DestroyRef VarPtr(cstObject)
   End If
   
   SystemTrayWndProc = DefWindowProc(hWnd, uMsg, wParam, lParam)

End Function

Public Sub DestroyRef(ByVal nObject As Long)

Dim lngValue As Long

   Call CopyMemory(ByVal nObject, ByVal VarPtr(lngValue), 4)

End Sub

Public Sub InitMessage()

   m_TaskbarRestart = RegisterWindowMessage("TaskbarCreated")

End Sub

Private Function DeRef(ByVal nPointer As Long) As clsSystemTray

   Call CopyMemory(ByVal VarPtr(DeRef), ByVal VarPtr(nPointer), 4)

End Function

