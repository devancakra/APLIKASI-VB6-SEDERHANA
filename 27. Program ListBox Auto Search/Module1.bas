Attribute VB_Name = "Module1"
Declare Function sendmessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As _
Long, ByVal wHsg As Long, ByVal wParam As Integer, ByVal lparam As Any) As Long
Public Const LB_FINDSTRING = &H18F

