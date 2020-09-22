VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScreenCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Function GetVersionEx& Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO)

Private Const VK_MENU = &H12
Private Const VK_SNAPSHOT = &H2C
Private Const KEYEVENTF_KEYUP = &H2

' used for dwPlatformId
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFO   '  148 Bytes
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type


Public Function SaveScreenToFile(ByVal strFile As String, Optional EntireScreen As Boolean = True) As Boolean

Dim altscan%
Dim snapparam%
Dim ret&, IsWin95 As Boolean
Dim verInfo As OSVERSIONINFO

  On Error GoTo errHand
    
  'Check if the File Exist
  If Dir(strFile) <> "" Then
    Kill strFile
    'Exit Function
  End If
  
  altscan% = MapVirtualKey(VK_MENU, 0)
  If EntireScreen = False Then
    keybd_event VK_MENU, altscan, 0, 0
    ' It seems necessary to let this key get processed before
    ' taking the snapshot.
  End If
    
  verInfo.dwOSVersionInfoSize = 148
  ret = GetVersionEx(verInfo)
  If verInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
    IsWin95 = True
  Else
    IsWin95 = False
  End If
    
  If EntireScreen = True And IsWin95 Then snapparam = 1
    
  DoEvents    ' These seem necessary to make it reliable
    
  ' Take the snapshot
  keybd_event VK_SNAPSHOT, snapparam, 0, 0
    
  DoEvents
    
  If EntireScreen = False Then keybd_event VK_MENU, altscan, KEYEVENTF_KEYUP, 0
    
  SavePicture Clipboard.GetData(vbCFBitmap), strFile
  
  SaveScreenToFile = True
  
  Exit Function
  
errHand:

  'Error handling
  SaveScreenToFile = False
  
End Function



