<div align="center">

## Screen Capture Class Module


</div>

### Description

This module will allow you to easily save screen captures. You can specify wether you want to capture the entire screen or just the active window. I've included a copy of the class module for download (since PSC doesn't do that good of a job at formating the code). Any comments or suggestions are welcome.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-07-02 09:56:10
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Screen Cap22061722001\.zip](https://github.com/Planet-Source-Code/screen-capture-class-module__1-24651/archive/master.zip)





### Source Code

<p>Option Explicit<br>
<br>
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As
Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)<br>
Private Declare Function MapVirtualKey Lib "user32" Alias
"MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long<br>
Private Declare Function GetVersionEx& Lib "kernel32" Alias
"GetVersionExA" (lpVersionInformation As OSVERSIONINFO)<br>
<br>
Private Const VK_MENU = &H12<br>
Private Const VK_SNAPSHOT = &H2C<br>
Private Const KEYEVENTF_KEYUP = &H2<br>
<br>
' used for dwPlatformId<br>
Private Const VER_PLATFORM_WIN32s = 0<br>
Private Const VER_PLATFORM_WIN32_WINDOWS = 1<br>
Private Const VER_PLATFORM_WIN32_NT = 2<br>
<br>
Private Type OSVERSIONINFO ' 148 Bytes<br>
dwOSVersionInfoSize As Long<br>
dwMajorVersion As Long<br>
dwMinorVersion As Long<br>
dwBuildNumber As Long<br>
dwPlatformId As Long<br>
szCSDVersion As String * 128<br>
End Type<br>
<br>
<br>
Public Function SaveScreenToFile(ByVal strFile As String, Optional EntireScreen As Boolean
= True) As Boolean</p>
<p><br>
Dim altscan%<br>
Dim snapparam%<br>
Dim ret&, IsWin95 As Boolean<br>
Dim verInfo As OSVERSIONINFO</p>
<blockquote>
 <p><br>
 On Error GoTo errHand<br>
 <br>
 'Check if the File Exist<br>
 If Dir(strFile) <> "" Then<br>
 Kill strFile<br>
 'Exit Function<br>
 End If<br>
 <br>
 altscan% = MapVirtualKey(VK_MENU, 0)<br>
 If EntireScreen = False Then<br>
 keybd_event VK_MENU, altscan, 0, 0<br>
 ' It seems necessary to let this key get processed before<br>
 ' taking the snapshot.<br>
 End If<br>
 <br>
 verInfo.dwOSVersionInfoSize = 148<br>
 ret = GetVersionEx(verInfo)<br>
 If verInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then<br>
 IsWin95 = True<br>
 Else<br>
 IsWin95 = False<br>
 End If<br>
 <br>
 If EntireScreen = True And IsWin95 Then snapparam = 1<br>
 <br>
 DoEvents ' These seem necessary to make it reliable<br>
 <br>
 ' Take the snapshot<br>
 keybd_event VK_SNAPSHOT, snapparam, 0, 0<br>
 <br>
 DoEvents<br>
 <br>
 If EntireScreen = False Then keybd_event VK_MENU, altscan, KEYEVENTF_KEYUP, 0<br>
 <br>
 SavePicture Clipboard.GetData(vbCFBitmap), strFile<br>
 <br>
 SaveScreenToFile = True<br>
 <br>
 Exit Function<br>
 <br>
 errHand:<br>
 <br>
 'Error handling<br>
 SaveScreenToFile = False</p>
</blockquote>
<p><br>
End Function<br>
<br>

