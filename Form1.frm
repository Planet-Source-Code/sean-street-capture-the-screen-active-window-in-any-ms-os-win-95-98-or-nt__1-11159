VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   14280
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   6525
      Left            =   120
      ScaleHeight     =   6465
      ScaleWidth      =   10245
      TabIndex        =   1
      Top             =   300
      Width           =   10305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   915
      Left            =   11880
      TabIndex        =   0
      Top             =   6300
      Width           =   2085
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Picture1.Picture = LoadPicture()
    CaptureScreen True
    Picture1.Picture = Clipboard.GetData(vbCFBitmap)
    
End Sub


Private Sub CaptureScreen(blnActiveWindow As Boolean)
    Dim intALTScan As Integer
    Dim intMSMistake As Integer
    
    intMSMistake = 0
    
    intALTScan = MapVirtualKey(VK_MENU, 0)  'returns the scan code key for the ALT button
    
    'if we want to capture just the active window
    'and not the whole screen
    If blnActiveWindow Then keybd_event VK_MENU, intALTScan, 0, 0

    'The documentation for the keybd_event with relation to
    'Windows 95 is backwards.  So if our operating system is not
    'Windows NT (if it is Win 95/98) then we have to switch the
    'variables

    If blnActiveWindow And Not (IsWindowsNT) Then intMSMistake = 1

    DoEvents    'stabalizes the system as it captures the data
    
    'captures the data
    keybd_event VK_SNAPSHOT, intMSMistake, 0, 0
    
    DoEvents
    
    'deactivates the ALT button press event
    If blnActiveWindow Then keybd_event VK_MENU, intALTScan, KEYEVENTF_KEYUP, 0

End Sub

Private Function IsWindowsNT() As Boolean
    Dim lngHolder As Long
    
    typSysInfo.dwOSVersionInfoSize = 148
    lngHolder = GetVersionEx(typSysInfo)
    
    IsWindowsNT = (typSysInfo.dwPlatformId = 2)

End Function
