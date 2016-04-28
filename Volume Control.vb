Option Explicit

Const VK_VOLUME_MUTE = &HAD 'Windows 2000/XP: Volume Mute key
Const VK_VOLUME_DOWN = &HAE  'Windows 2000/XP: Volume Down key
Const VK_VOLUME_UP = &HAF  'Windows 2000/XP: Volume Up key

Private Declare Sub keybd_event Lib "user32" ( _
   ByVal bVk As Byte, ByVal bScan As Byte, _
   ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Sub VolUp()
   keybd_event VK_VOLUME_UP, 0, 1, 0
   keybd_event VK_VOLUME_UP, 0, 3, 0
End Sub
Sub VolDown()
   keybd_event VK_VOLUME_DOWN, 0, 1, 0
   keybd_event VK_VOLUME_DOWN, 0, 3, 0
End Sub

Sub VolToggle()
   keybd_event VK_VOLUME_MUTE, 0, 1, 0
End Sub
