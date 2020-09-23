VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Kill AOL Ads!"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "adkill.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)

Const WM_CLOSE = &H10
Const GW_CHILD = 5
Const GW_HWNDFIRST = 0
Const GW_HWNDLAST = 1
Const GW_HWNDNEXT = 2
Const GW_HWNDPREV = 3
Function GetClass(hwnd As Long)
Temp$ = String$(250, 0)
TheClass% = GetClassName(hwnd, Temp$, 250)

GetClass = Temp$
End Function
Function FindWindowByClass(TheParent As Long, TheClass As String)
Temp& = GetWindow(TheParent, 5)
If UCase(Mid(GetClass(Temp&), 1, Len(TheClass))) Like UCase(TheClass) Then GoTo Found
Temp& = GetWindow(TheParent, GW_CHILD)
If UCase(Mid(GetClass(Temp&), 1, Len(TheClass))) Like UCase(TheClass) Then GoTo Found

While Temp&
Temp2& = GetWindow(TheParent, 5)
If UCase(Mid(GetClass(Temp2&), 1, Len(TheClass))) Like UCase(TheClass) Then GoTo Found
Temp& = GetWindow(Temp&, 2)
If UCase(Mid(GetClass(Temp&), 1, Len(TheClass))) Like UCase(TheClass) Then GoTo Found
Wend
FindChildByClass = 0

Found:
FindWindowByClass = Temp&

End Function


Private Sub Form_Load()
Me.Hide
End Sub

Private Sub Timer1_Timer()
DoEvents 'Put doevents so the timer doesn't keep triggering while something's going on
Dim AOL As Long
AOL = FindWindow("AOL Frame25", vbNullString) 'Find AOL Window handle
MDI% = FindWindowByClass(AOL, "MDIClient") 'Find handle of main AOL screen
If MDI% = 0 Then Exit Sub 'If AOL wasn't found, avoid using unnecesary CPU cycles
Window& = GetWindow(MDI%, GW_CHILD) 'Get the first child window
Do Until Window& = 0 'Find all child windows
'Normally, we'd search 5 or so levels down...
'But since AOL has it's own images and very few ads at that level
'We'll do a superficial search(1 level down)
TheWindow = FindWindowByClass(Window&, "_AOL_Image") 'Find any ad-like windows by their class
If TheWindow <> 0 Then Call SendMessage(TheWindow, WM_CLOSE, 0, 0) 'Close any ads
'This may interfere with other AOL images also. I use Compuserve
'So I don't see many images aside from ads. But it may pose a
'Problem for AOL users
Window& = GetWindow(Window&, GW_HWNDNEXT) 'Go to the next child window
Loop
End Sub
