Attribute VB_Name = "Module1"
Option Explicit

'-- You Won't find many comments in here because I just decided to post it for the hell of it
'-- If you can't understand it...well, here's a good time to learn VB

Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const MB_OK = &H0&
Public Const MB_ICONASTERISK = &H40&

Public Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                           ByVal hWnd As Long, _
                           ByVal wMsg As Long, _
                           ByVal wParam As Long, _
                           lParam As Any) As Long

Public gblnFormLoaded As Boolean
