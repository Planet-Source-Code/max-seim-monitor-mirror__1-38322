VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   Caption         =   "Monitor Mirror"
   ClientHeight    =   2625
   ClientLeft      =   2805
   ClientTop       =   2805
   ClientWidth     =   3450
   ControlBox      =   0   'False
   Icon            =   "mirror.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3450
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "Minimize"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "Close -- Save Settings"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'    Monitor Mirror
'
'    By:  Max Seim
'         mlseim@mmm.com
'
'         August 26, 2002
'
'
'    Uses a couple of useful routines:
'    1) Utilize Registry to retain form position and size
'    2) Utilize the "stay on top", so form is always on top
'
'
Const HWND_TOPMOST& = -1
Const SWP_NOMOVE& = &H2&
Const SWP_NOSIZE& = &H1&
Private Declare Function SetWindowPos& Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Function StayOnTop(Form As Form)
' This function sets the form to always stay on top.
Dim lFlags As Long
Dim lStay As Long
lFlags = SWP_NOSIZE Or SWP_NOMOVE
lStay = SetWindowPos(Form.hWnd, HWND_TOPMOST, 0, 0, 0, 0, lFlags)
End Function

Private Sub Form_Load()
'Get settings from registry
'For first-time run, the settings are not set. Use a default ex. 0
Form1.Top = GetSetting("MyApp", "FormSettings", "FormTop", 0) '0 = Default
Form1.Left = GetSetting("MyApp", "FormSettings", "FormLeft", 0) '0 = Default
Form1.Width = GetSetting("MyApp", "FormSettings", "FormWidth", 3500) '3500 = Default
Form1.Height = GetSetting("MyApp", "FormSettings", "FormHeight", 2600) '2600 = Default
Call StayOnTop(Me)
End Sub

Private Sub Command1_Click()
'Save the settings for form1.top and form1.left
'save keys under : HKEY_CURRENT_USER\Software\VB and VBA Program Settings
'Run regedit.exe and check out HKEY_CURRENT_USER\Software\VB and VBA Program Settings
SaveSetting "MyApp", "FormSettings", "FormTop", Form1.Top
SaveSetting "MyApp", "FormSettings", "FormLeft", Form1.Left
SaveSetting "MyApp", "FormSettings", "FormWidth", Form1.Width
SaveSetting "MyApp", "FormSettings", "FormHeight", Form1.Height
Unload Me
End Sub

Private Sub Command2_Click()
'Minimize Window
Form1.WindowState = 1
End Sub
