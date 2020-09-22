VERSION 5.00
Begin VB.Form frmSender 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sender"
   ClientHeight    =   1152
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3048
   Icon            =   "frmSender.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1152
   ScaleWidth      =   3048
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   372
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   1212
   End
   Begin VB.TextBox txtText 
      Height          =   288
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2772
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "S&end"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1212
   End
End
Attribute VB_Name = "frmSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type CopyDataStruct
   dwData As Long
   cbData As Long
   lpData As Long
End Type

Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Sub SendCommand(ByVal SendData As String)

Const WM_ACTIVATE       As Long = &H6
Const WM_COPYDATA       As Long = &H4A
Const APP_NAME          As String = "Receiver"
Const CLASS_NAME_HIDDEN As String = "SystemTray_HiddenWindow_"
Const RECEIVED_DATA     As String = "Received_Data_"

Dim bytBuffer(1 To 255) As Byte
Dim cdsData             As CopyDataStruct
Dim lngWindow           As Long
Dim strData             As String

   If Len(RECEIVED_DATA & SendData) > 254 Then
      MsgBox "Data to long!  (Max. 238 characters!)"
      Exit Sub
   End If
   
   lngWindow = FindWindow(vbNullString, CLASS_NAME_HIDDEN & APP_NAME)
   strData = RECEIVED_DATA & SendData
   
   If lngWindow = 0 Then lngWindow = FindWindow(vbNullString, APP_NAME)
   If lngWindow = 0 Then Exit Sub
   
   Call CopyMemory(bytBuffer(1), ByVal strData, Len(strData))
   
   With cdsData
      .dwData = 3
      .cbData = Len(strData) + 1
      .lpData = VarPtr(bytBuffer(1))
   End With
   
   SendMessage lngWindow, WM_COPYDATA, WM_ACTIVATE, cdsData

End Sub
     
Private Sub cmdSend_Click()

   If Trim(txtText.Text) = "" Then Exit Sub
   
   Call SendCommand(txtText.Text)

End Sub

Private Sub cmdStop_Click()

   Call SendCommand(Chr(1) & "QUIT")
   
   Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Unload Me

End Sub
