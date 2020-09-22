Attribute VB_Name = "mdlHook"
Option Explicit

' Private Constant
Private Const GWL_WNDPROC As Long = -4

' Private Variable
Private PrevSystemTray    As Long

' Private API's
Private Declare Function StrLen Lib "Kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Sub HookSystemTray(ByVal hWnd As Long)

   If PrevSystemTray Then
      SetWindowLong hWnd, GWL_WNDPROC, PrevSystemTray
      PrevSystemTray = 0
      
   Else
      PrevSystemTray = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SystemTrayWinProc)
   End If

End Sub

Private Function SystemTrayWinProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim bytBuffer(1 To 255) As Byte
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
               
               Call frmReceiver.ReceivedData(strData)
            End If
         End If
      End With
   End If
   
   SystemTrayWinProc = CallWindowProc(PrevSystemTray, hWnd, uMsg, wParam, lParam)

End Function
