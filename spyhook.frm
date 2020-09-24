VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Cross Task subclassing -using system wide hooks-venky_dude"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Start Notepad && Hook IT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "This example is similar to the cross task subclasing demo in the Spyworks product of desaware."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   6135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Intro"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Instructions"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   $"spyhook.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   $"spyhook.frx":00DC
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Const MF_STRING = &H0&
Private Const WH_CALLWNDPROC = 4
Dim n As Long
Private Sub Command1_Click()
Shell "notepad.exe", vbNormalFocus
Dim hdl As Long
Dim s As String
hdl = FindWindow(s, "Untitled - NotePad")
Dim ml As Long
ml = GetMenu(hdl)
Dim msl As Long
msl = GetSubMenu(ml, 1)
Dim res As Long
res = AppendMenu(msl, MF_STRING, 35, "Hooked by venky!!!")
Dim hnd As Long
hnd = LoadLibrary("keybuster.dll")
Dim hnd2 As Long
hnd2 = GetProcAddress(hnd, "CallWndProc")
n = SetWindowsHookEx(WH_CALLWNDPROC, hnd2, hnd, 0)
End Sub



Private Sub Form_Unload(Cancel As Integer)
Dim x As Long
x = UnhookWindowsHookEx(n)
End Sub
