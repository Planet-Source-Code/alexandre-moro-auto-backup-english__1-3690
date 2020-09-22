VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto Backup"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   59000
      Left            =   0
      Top             =   4920
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Minimize and save changes (ESC)"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   4800
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   8070
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Hour/Day"
      TabPicture(0)   =   "Main.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "CheckBox4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CheckBox1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "OptionButton2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CheckBox2(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CheckBox2(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "CheckBox2(6)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CheckBox2(5)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "CheckBox2(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "CheckBox2(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "CheckBox2(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "OptionButton3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label8"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "OptionButton1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label7"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label3"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "MaskEdBox2"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "MaskEdBox1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Command6"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Files/Directories"
      TabPicture(1)   =   "Main.frx":0E5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "File1"
      Tab(1).Control(1)=   "Command4"
      Tab(1).Control(2)=   "List1"
      Tab(1).Control(3)=   "Command1"
      Tab(1).Control(4)=   "Dir1"
      Tab(1).Control(5)=   "Drive1"
      Tab(1).Control(6)=   "Command5"
      Tab(1).Control(7)=   "Label1"
      Tab(1).Control(8)=   "CheckBox3"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Destiny"
      TabPicture(2)   =   "Main.frx":0E7A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Drive2"
      Tab(2).Control(1)=   "Dir2"
      Tab(2).Control(2)=   "Text1"
      Tab(2).Control(3)=   "Label2"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Progress"
      TabPicture(3)   =   "Main.frx":0E96
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command7"
      Tab(3).Control(1)=   "Label14"
      Tab(3).Control(2)=   "Label13"
      Tab(3).Control(3)=   "Label12"
      Tab(3).Control(4)=   "Label11"
      Tab(3).Control(5)=   "Label10"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "About"
      TabPicture(4)   =   "Main.frx":0EB2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label16"
      Tab(4).Control(1)=   "Label15"
      Tab(4).ControlCount=   2
      Begin VB.CommandButton Command7 
         Caption         =   "See log file..."
         Height          =   495
         Left            =   -73800
         TabIndex        =   43
         Top             =   3720
         Width           =   3255
      End
      Begin VB.FileListBox File1 
         Height          =   1650
         Left            =   -72000
         MultiSelect     =   2  'Extended
         TabIndex        =   15
         Top             =   2040
         Width           =   2535
      End
      Begin VB.CommandButton Command4 
         Caption         =   "+ Directory"
         Height          =   375
         Left            =   -73440
         TabIndex        =   14
         Top             =   3885
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Height          =   1230
         ItemData        =   "Main.frx":0ECE
         Left            =   -74880
         List            =   "Main.frx":0ED0
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   5415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exclude"
         Height          =   375
         Left            =   -70800
         TabIndex        =   12
         Top             =   3885
         Width           =   1215
      End
      Begin VB.DirListBox Dir1 
         Height          =   1215
         Left            =   -74880
         TabIndex        =   11
         Top             =   2460
         Width           =   2775
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   -74880
         TabIndex        =   10
         Top             =   2040
         Width           =   2775
      End
      Begin VB.CommandButton Command5 
         Caption         =   "+ File(s)"
         Height          =   375
         Left            =   -72120
         TabIndex        =   9
         Top             =   3885
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Do backup now!"
         Height          =   495
         Left            =   3720
         TabIndex        =   8
         Top             =   3780
         Width           =   1695
      End
      Begin VB.DriveListBox Drive2 
         Height          =   315
         Left            =   -73680
         TabIndex        =   5
         Top             =   1200
         Width           =   2895
      End
      Begin VB.DirListBox Dir2 
         Height          =   2565
         Left            =   -73680
         TabIndex        =   4
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   -74640
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   4935
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   285
         Left            =   1845
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   930
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   285
         Left            =   1845
         TabIndex        =   7
         Top             =   1290
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Please mail me your comments and suggestions!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74280
         MouseIcon       =   "Main.frx":0ED2
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Top             =   2640
         Width           =   4125
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "©1999 Alexandre Moro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73200
         TabIndex        =   44
         Top             =   1800
         Width           =   2010
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   -74880
         TabIndex        =   42
         Top             =   2880
         Width           =   5400
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   -74880
         TabIndex        =   41
         Top             =   2280
         Width           =   5400
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74880
         TabIndex        =   40
         Top             =   1800
         Width           =   5400
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   -74880
         TabIndex        =   39
         Top             =   1200
         Width           =   5400
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74880
         TabIndex        =   38
         Top             =   720
         Width           =   5400
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Files / Directories to be copied:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   37
         Top             =   480
         Width           =   2205
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destiny directory:"
         Height          =   195
         Left            =   -74640
         TabIndex        =   36
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Backup when..."
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "hours:minutes"
         Height          =   195
         Left            =   2640
         TabIndex        =   34
         Top             =   1335
         Width           =   975
      End
      Begin MSForms.OptionButton OptionButton1 
         Height          =   345
         Left            =   1560
         TabIndex        =   33
         Top             =   900
         Width           =   405
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "714;609"
         Value           =   "1"
         GroupName       =   "a"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "hours:minutes"
         Height          =   195
         Left            =   2640
         TabIndex        =   32
         Top             =   975
         Width           =   975
      End
      Begin MSForms.OptionButton OptionButton3 
         Height          =   345
         Left            =   1560
         TabIndex        =   31
         Top             =   2100
         Width           =   750
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "1323;609"
         Value           =   "1"
         Caption         =   "Daily"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox CheckBox2 
         Height          =   345
         Index           =   2
         Left            =   1560
         TabIndex        =   30
         Top             =   2460
         Width           =   975
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1720;609"
         Value           =   "0"
         Caption         =   "Monday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox CheckBox2 
         Height          =   345
         Index           =   3
         Left            =   2880
         TabIndex        =   29
         Top             =   2460
         Width           =   1020
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1799;609"
         Value           =   "0"
         Caption         =   "Tuesday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox CheckBox2 
         Height          =   345
         Index           =   4
         Left            =   4080
         TabIndex        =   28
         Top             =   2460
         Width           =   1260
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2222;609"
         Value           =   "0"
         Caption         =   "Wednesday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox CheckBox2 
         Height          =   345
         Index           =   5
         Left            =   1560
         TabIndex        =   27
         Top             =   2820
         Width           =   1065
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1879;609"
         Value           =   "0"
         Caption         =   "Thursday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox CheckBox2 
         Height          =   345
         Index           =   6
         Left            =   2880
         TabIndex        =   26
         Top             =   2820
         Width           =   825
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1455;609"
         Value           =   "0"
         Caption         =   "Friday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox CheckBox2 
         Height          =   345
         Index           =   7
         Left            =   4080
         TabIndex        =   25
         Top             =   2820
         Width           =   1035
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1826;609"
         Value           =   "0"
         Caption         =   "Saturday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox CheckBox2 
         Height          =   345
         Index           =   1
         Left            =   1560
         TabIndex        =   24
         Top             =   3180
         Width           =   945
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1667;609"
         Value           =   "0"
         Caption         =   "Sunday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "When:"
         Height          =   195
         Left            =   960
         TabIndex        =   23
         Top             =   2175
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Always at:"
         Height          =   195
         Left            =   720
         TabIndex        =   22
         Top             =   975
         Width           =   720
      End
      Begin MSForms.OptionButton OptionButton2 
         Height          =   345
         Left            =   1560
         TabIndex        =   21
         Top             =   1260
         Width           =   405
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "714;609"
         Value           =   "0"
         GroupName       =   "a"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Or each:"
         Height          =   195
         Left            =   825
         TabIndex        =   20
         Top             =   1335
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "(interval initiated from now, or when the application starts)"
         Height          =   390
         Left            =   1560
         TabIndex        =   19
         Top             =   1620
         Width           =   2580
         WordWrap        =   -1  'True
      End
      Begin MSForms.CheckBox CheckBox1 
         Height          =   345
         Left            =   360
         TabIndex        =   18
         Top             =   3540
         Width           =   1275
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2249;609"
         Value           =   "1"
         Caption         =   "Save log file"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox CheckBox3 
         Height          =   495
         Left            =   -74760
         TabIndex        =   17
         Top             =   3825
         Width           =   1215
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2143;873"
         Value           =   "1"
         Caption         =   "Include Subdirs"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox CheckBox4 
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   3960
         Width           =   2055
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "3625;661"
         Value           =   "0"
         Caption         =   "Incremental backup"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Menu mnu_1 
      Caption         =   "mnu_1"
      Visible         =   0   'False
      Begin VB.Menu MnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu MnuBackup 
         Caption         =   "Backup now!"
      End
      Begin VB.Menu MnuQuit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'********** Auto Backup ***********
'******* ©1999 Alexandre Moro *******
'You can freely distribute this source code,
'   but if you do any modification please
'                 let me know!
'
'       Comments and suggestions:
'         alb@cwb.matrix.com.br



Dim NLoops As Integer, LoopDup As Integer, ListWithFocus As Boolean, Days As Byte
Dim sRet As String, Ret As Long, MskErr1 As Boolean, MskErr2 As Boolean
Dim DestinyDir As String, NoIniArchive As Boolean
Dim WindowsDir As String, NLoopsTimer As Byte, Interval As Date, IniTime As Date
Dim Default As Boolean, LastBackup As Date, Result As Long, Msg As Long, OpenError As Boolean
Dim XDir(2) As New Collection, FromPath As String

Private Const Arq = "Autobak.ini"
Private Const SW_SHOW = 5

Private Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String)
Private Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String)
Private Declare Function GetWindowsDirectory& Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
    
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private nid As NOTIFYICONDATA

Private Type ListaArqs
    Nome As String
    Tamanho As Long
End Type

Private Files() As ListaArqs
Private Sub GetDirs(Path As String)

    'on error Resume Next
    Dim vDirName As String, LastDir As String
    Dim i As Integer
    
    'Adjust so No Deletion of Drive
    If Len(Path$) < 4 Then Exit Sub

    If Right(Path$, 1) <> "\" Then
        XDir(0).Add Path$
        Path$ = Path$ & "\"
    End If

    vDirName = Dir(Path, vbDirectory) ' Retrieve the first entry.

    Do While vDirName <> ""
        If vDirName <> "." And vDirName <> ".." Then
            If (GetAttr(Path & vDirName)) = vbDirectory Then
                LastDir = vDirName
                'Finds Directory Name then Repeats
                GetDirs (Path$ & vDirName)
                vDirName = Dir(Path$, vbDirectory)

                Do Until vDirName = LastDir Or vDirName = ""
                    vDirName = Dir
                Loop

                If vDirName = "" Then Exit Do
            End If
        End If
    
    vDirName = Dir
    
    Loop

End Sub

Private Function ExtractText(FullText As String, token As String, Optional StartAtLeft = True, Optional IncludeLeftSide = True) As String
'ExtractText(Path$, ":", False, False)
    
    Dim i As Integer
    If StartAtLeft = True And IncludeLeftSide = True Then
        ExtractText = FullText
        For i = 1 To Len(FullText)
            If Mid(FullText, i, 1) = token Then
                ExtractText = Left(FullText, i - 1)
                Exit Function
            End If
        Next

    ElseIf StartAtLeft = True And IncludeLeftSide = False Then
        ExtractText = FullText
        For i = 1 To Len(FullText)
            If Mid(FullText, i, 1) = token Then
                ExtractText = Right(FullText, Len(FullText) - i)
                Exit Function
            End If
        Next
    
    ElseIf StartAtLeft = False And IncludeLeftSide = True Then
        ExtractText = ""
        For i = Len(FullText) To 1 Step -1
            If Mid(FullText, i, 1) = token Then
                ExtractText = Left(FullText, i - 1)
                Exit Function
            End If
        Next

    ElseIf StartAtLeft = False And IncludeLeftSide = False Then
        ExtractText = ""
        For i = Len(FullText) To 1 Step -1
            If Mid(FullText, i, 1) = token Then
                ExtractText = Right(FullText, Len(FullText) - i)
                Exit Function
            End If
        Next
    End If

End Function


Private Sub MtxAdicionaArq(CamCompleto As String)
    
    If UBound(Files) = 1 Then
        Files(1).Nome = CamCompleto
        Files(1).Tamanho = FileLen(CamCompleto)
        ReDim Preserve Files(2)
    Else
        Files(UBound(Files)).Nome = CamCompleto
        Files(UBound(Files)).Tamanho = FileLen(CamCompleto)
        ReDim Preserve Files(UBound(Files) + 1)
    End If

End Sub

Private Sub MtxAdicionaDir(ByVal Caminho As String)
On Error GoTo erro

    Dim B As String, n As Integer, ShortPath As String
    
    If Not Right(Caminho, 1) = "*" Then Caminho = Caminho & "*.*"

    ShortPath = Left(Caminho, Len(Caminho) - 3)
            
    If Not UBound(Files) = 1 Then
        n = UBound(Files) + 1
        ReDim Preserve Files(n)
    End If
    
    B = Dir(Caminho)
    If B = "" Then
        Exit Sub
    Else
        Files(UBound(Files) - 1).Nome = ShortPath & B
        Files(UBound(Files) - 1).Tamanho = FileLen(ShortPath & B)
    End If

    Do
    B = Dir
    If B = "" Then Exit Do
        
    With Files(n)
        .Nome = ShortPath & B
        .Tamanho = FileLen(ShortPath & B)
    End With
    n = n + 1
    ReDim Preserve Files(n)
    Loop

Saída:
    Exit Sub
    
erro:
    MsgBox "MtxAddDir:" & vbLf & vbLf & Err.Number & ":" & Err.Description, vbCritical
    Resume Saída

End Sub

Private Sub AddItem(OnlyFile As Boolean, Optional WithSubs As Boolean = False)
On Error GoTo erro

    Screen.MousePointer = vbHourglass

    Dim AddPath As String
    
    If Right(Dir1.Path, 1) = "\" Then
        AddPath = Dir1.Path
    Else
        AddPath = Dir1.Path & "\"
    End If
    
    If Not OnlyFile Then
        
        If WithSubs Then
            Dim i As Integer, d As String
            GetDirs (AddPath)
            For i = 1 To XDir(0).Count
                If VerificaDup(XDir(0).Item(i) & "\*.*") Then
                    MsgBox "This item is already on the list:" & vbLf & vbLf & XDir(0).Item(i) & "\*.*", vbExclamation
                Else
                    List1.AddItem XDir(0).Item(i) & "\*.*"
                End If
            Next i
            For i = XDir(0).Count To 1 Step -1
                XDir(0).Remove (i)
            Next i
        End If
        
        If List1.ListCount = 0 Then
            List1.AddItem AddPath & "*.*"
            GoTo Saída
        Else
            If VerificaDup(AddPath & "*.*") Then
                MsgBox "This item is already on the list:" & vbLf & vbLf & AddPath & "*.*", vbExclamation
                GoTo Saída
            Else
                List1.AddItem AddPath & "*.*"
                GoTo Saída
            End If
        End If
        
    Else
    
        Dim Entries As Integer
        For NLoops = 0 To File1.ListCount - 1
            If File1.Selected(NLoops) Then
                Entries = Entries + 1
                If Entries > 1 Then GoTo cont
            End If
        Next NLoops

cont:
        If Entries = 1 Then
            If VerificaDup(AddPath & File1.FileName) Then
                MsgBox "This item is already on the list:" & vbLf & vbLf & AddPath & File1.FileName, vbExclamation
                GoTo Saída
            Else
                List1.AddItem AddPath & File1.FileName
                GoTo Saída
            End If
        ElseIf Entries > 1 Then
            For NLoops = 0 To File1.ListCount - 1
                If File1.Selected(NLoops) Then
                    If VerificaDup(AddPath & File1.List(NLoops)) Then
                        MsgBox "This item is already on the list:" & vbLf & vbLf & AddPath & File1.List(NLoops), vbExclamation
                    Else
                        List1.AddItem AddPath & File1.List(NLoops)
                    End If
                End If
            Next NLoops
        End If
        
    End If
    
Saída:
    Screen.MousePointer = vbDefault
    Exit Sub
    
erro:
    MsgBox Err.Number & vbLf & Err.Description, vbCritical
    Resume Saída
                    
End Sub

Private Sub Backup()
On Error GoTo erro

    Screen.MousePointer = vbHourglass
    
    Dim DateBak As Date, TimeBak As Date, ErrString As String
    Dim NDirs As Integer, File As String, TskID As Double, TotFiles As Long, TotalFilesCopied As Long
    Dim ErroDest As Byte, ArqAtr As Byte, Tam As Long
        
    SSTab1.Tab = 3
    
    TimeBak = Now
    DateBak = Date
    
    Me.Caption = "Creating file list..."
    
    If Not Right(DestinyDir, 1) = "\" Then DestinyDir = DestinyDir & "\"

    For NLoops = 0 To List1.ListCount - 1
        If Right(List1.List(NLoops), 1) = "*" Then
            MtxAdicionaDir (Left(List1.List(NLoops), Len(List1.List(NLoops)) - 3))
        Else
            MtxAdicionaArq (List1.List(NLoops))
        End If
    Next NLoops

    Me.Caption = "Doing the backup..."
    If CheckBox1 Then
        Open WindowsDir & "Log Autobak.txt" For Output As #1
        Print #1, "Initializing backup at " & Now
        Print #1,
    End If
    
    Label10.Caption = "Copying now"
    Label12.Caption = "to"
    
    TotFiles = UBound(Files) - 1
    For NLoops = 0 To TotFiles
        DoEvents
        If Not Files(NLoops).Nome = "" Then
            ArqAtr = GetAttr(Files(NLoops).Nome)
            Label11.Caption = Files(NLoops).Nome
            Label13.Caption = DestinyDir & ReturnFileName(Files(NLoops).Nome)
            Label14.Caption = "File " & NLoops & " of " & TotFiles
cont:
            If CheckBox4 Then
                If ArqAtr And vbArchive <> 0 Then
                    If CheckBox1 Then Print #1, Files(NLoops).Nome & " --> " & DestinyDir & ReturnFileName(Files(NLoops).Nome) & ", status: ";
                    FileCopy Files(NLoops).Nome, DestinyDir & ReturnFileName(Files(NLoops).Nome)
                    SetAttr Files(NLoops).Nome, (ArqAtr - vbArchive)
                    If CheckBox1 Then Print #1, "Ok!"
                    Tam = Tam + FileLen(Files(NLoops).Nome)
                    TotalFilesCopied = TotalFilesCopied + 1
                End If
            Else
                If CheckBox1 Then Print #1, Files(NLoops).Nome & " --> " & DestinyDir & ReturnFileName(Files(NLoops).Nome) & ", status: ";
                FileCopy Files(NLoops).Nome, DestinyDir & ReturnFileName(Files(NLoops).Nome)
                If CheckBox1 Then Print #1, "Ok!"
                Tam = Tam + FileLen(Files(NLoops).Nome)
                TotalFilesCopied = TotalFilesCopied + 1
            End If
            Label14.Caption = "File " & NLoops & " of " & TotFiles & ", total: " & _
                        Format(Tam / 1024 / 1024, "standard") & " Mb"
        End If
    Next NLoops

Saída:
    If CheckBox1 Then
        Print #1,
        Print #1, "Copied " & TotalFilesCopied & " files, " & Format(Tam / 1024 / 1024, "standard") & " Mb, from " & _
            Format(TimeBak, "short time") & " to " & Format(Time, "short time") & " of " & _
            Format(DateBak, "short date") & "."
        Close #1
    End If
    Label10.Caption = ""
    Label11.Caption = ""
    Label12.Caption = ""
    Label13.Caption = ""
    Label14.Caption = "Copied " & TotalFilesCopied & " files, " & Format(Tam / 1024 / 1024, "standard") & " Mb, from " & _
        Format(TimeBak, "short time") & " to " & Format(Time, "short time") & " of " & _
        Format(DateBak, "short date") & "."
    ReDim Files(0)
    Me.Caption = "Auto Backup"
    Screen.MousePointer = vbDefault
    Exit Sub

erro:
    ErrString = vbLf & vbLf & "While trying to copy:" & vbLf & Files(NLoops).Nome & _
        vbLf & "to" & vbLf & DestinyDir & ReturnFileName(Files(NLoops).Nome) & vbLf & _
        vbLf & "Try again?"
    
    If CheckBox1 Then Print #1, "ERROR: " & Err.Number & " - " & Err.Description;
    
    Select Case Err.Number
        
        Case 5      'Invalid procedure call ???
            Resume Next
                    
        Case 52    'Bad filename
            MsgBox "Bad filename! (erro 52)" & vbLf & vbLf & Files(NLoops).Nome, vbExclamation
            Resume Next
            
        Case 53     'File not found
            MsgBox "File not found! (erro 53)" & vbLf & vbLf & Files(NLoops).Nome, vbExclamation
            Resume Next
                    
        Case 57     'Device I/O error
            If MsgBox("Destiny disk not ready! (erro 57)" & ErrString, vbCritical + vbYesNo) = vbYes Then Resume cont
            
        Case 61     'Disk full
            If MsgBox("Destiny disk full! (error 61)" & ErrString, vbExclamation + vbYesNo) = vbYes Then Resume cont
                    
        Case 70    'Permission denied
            If MsgBox("Destiny directory or drive protected! (error 70)" & ErrString, vbCritical + vbYesNo) = vbYes Then Resume cont
            
        Case 71    'Disk not ready
            If MsgBox("Destiny disk not ready! (error 71)" & ErrString, vbCritical + vbYesNo) = vbYes Then Resume cont
                
        Case 75     'Path/file access error
            SetAttr DestinyDir & ReturnFileName(Files(NLoops).Nome), (GetAttr(DestinyDir & ReturnFileName(Files(NLoops).Nome)) - vbReadOnly)
            Resume cont
        
        Case 76     'Path not found
            If MsgBox("Destiny directory unavailable! (error 76)" & ErrString, vbCritical + vbYesNo) = vbYes Then Resume cont
        
        Case Else
            If MsgBox("PANIC!!" & vbLf & vbLf & Err.Number & ": " & Err.Description & ErrString, vbCritical + vbYesNo) = vbYes Then Resume cont
    
    End Select
    
    Resume Saída
    
End Sub

Private Function ReturnFileName(ByVal Arq As String) As String
'Arq is the full path, returns only the filename
    
    Dim n As Integer
    
    For n = Len(Arq) To 1 Step -1
        If Mid(Arq, n, 1) = "\" Then
            ReturnFileName = Right(Arq, Len(Arq) - n)
            Exit Function
        End If
    Next n

End Function
Private Sub CheckTime()
On Error GoTo erro

    If OptionButton1 And Not IniTime = vbEmpty Then
        If IniTime = TimeSerial(Hour(Time), Minute(Time), 0) Then
            Me.Caption = "Doing the Backup..."
            Me.Refresh
            Backup
            LastBackup = TimeSerial(Hour(Time), Minute(Time), 0)
            Me.Caption = "Auto Backup"
            Me.Refresh
        End If
    End If
    
    If OptionButton2 And Not Interval = vbEmpty Then
        If TimeSerial(Hour(Time), Minute(Time), 0) = TimeValue(Interval + LastBackup) Then
            Me.Caption = "Doing the Backup..."
            Me.Refresh
            Backup
            LastBackup = TimeSerial(Hour(Time), Minute(Time), 0)
            Me.Caption = "Auto Backup"
            Me.Refresh
        End If
    End If
    
Saída:
    Exit Sub
    
erro:
    If Not Err.Number = 13 Then MsgBox Err.Number & vbLf & Err.Description
    Resume Saída
    
End Sub

Private Sub Initialize()
On Error GoTo erro

    Dim Lenght As Byte
    
    WindowsDir = String(255, 0)
    Lenght = GetWindowsDirectory(WindowsDir, 254)
    WindowsDir = Left(WindowsDir, Lenght)
    
    If Not Right(WindowsDir, 1) = "\" Then WindowsDir = WindowsDir & "\"
    
    If Dir(WindowsDir & "Autobak.ini") = "" Then
        If Dir(WindowsDir & "Autobak.bak") <> "" Then
            FileCopy WindowsDir & "Autobak.bak", WindowsDir & "Autobak.ini"
        Else
            NoIniArchive = True
        End If
    End If
        
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("When", "AlwaysAt", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then
        If sRet = "???" Then
            IniTime = vbEmpty
        Else
            MaskEdBox1.Text = sRet
            IniTime = TimeSerial(Hour(MaskEdBox1.Text), Minute(MaskEdBox1.Text), 0)
        End If
    End If

    sRet = String(255, 0)
    Ret = GetPrivateProfileString("When", "Each", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then
        If sRet = "???" Then
            Interval = vbEmpty
        Else
            MaskEdBox2.Text = sRet
            Interval = TimeSerial(Hour(MaskEdBox2.Text), Minute(MaskEdBox2.Text), 0)
        End If
    End If
    
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("When", "Default", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then
        If sRet = "False" Then
            Default = False
        Else
            Default = True
        End If
    End If
    
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("When", "Days", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then
        Dim BsRet As Byte
        BsRet = CByte(sRet)
        If Int(BsRet / 64) = 1 Then CheckBox2(7).Value = True: BsRet = BsRet - 64
        If Int(BsRet / 32) = 1 Then CheckBox2(6).Value = True: BsRet = BsRet - 32
        If Int(BsRet / 16) = 1 Then CheckBox2(5).Value = True: BsRet = BsRet - 16
        If Int(BsRet / 8) = 1 Then CheckBox2(4).Value = True: BsRet = BsRet - 8
        If Int(BsRet / 4) = 1 Then CheckBox2(3).Value = True: BsRet = BsRet - 4
        If Int(BsRet / 2) = 1 Then CheckBox2(2).Value = True: BsRet = BsRet - 2
        If Int(BsRet / 1) = 1 Then CheckBox2(1).Value = True
    End If
    
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("Log", "Save", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then If sRet = "False" Then CheckBox1.Value = False
    
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("Backup", "Incremental", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then If sRet = "True" Then CheckBox4.Value = True
    
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("Destiny", "Dir", "", sRet, 255, Arq)
    sRet = Left(sRet, Ret)
    If Not Ret = 0 Then
        On Error GoTo erro1
        Dir2.Path = sRet
        Drive2.Drive = Left(sRet, 2)
        On Error GoTo erro
    End If
    
cont:
    DestinyDir = sRet
    Text1.Text = DestinyDir
    NLoops = 0
    ReDim Files(0)
    
    
start:
    sRet = String(255, 0)
    Ret = GetPrivateProfileString("Entries", NLoops, "", sRet, 255, Arq)
    If Ret = 0 Then LastBackup = TimeSerial(Hour(Time), Minute(Time), 0): Exit Sub
    sRet = Left(sRet, Ret)
    List1.AddItem sRet
    NLoops = NLoops + 1
    GoTo start

    
Saída:
    Exit Sub
    
erro:
    MsgBox Err.Number & vbLf & vbLf & Err.Description, vbCritical, "Initializing!"
    Resume Next
    
erro1:
    If Err.Number = 68 Or Err.Number = 76 Then
        'MsgBox "O diretório ou drive de destino não está disponível!" & vbLf & vbLf & _
            "Deixado como Default ""C:\""", vbExclamation
        'sRet = "C:\"
    Else
        MsgBox Err.Number & vbLf & Err.Description
    End If
    Resume cont
    
End Sub


Private Sub SaveChanges()
On Error GoTo erro

    Screen.MousePointer = vbHourglass
        
    On Error Resume Next
    Name WindowsDir & Arq As WindowsDir & "Autobak.bak"
    Kill WindowsDir & Arq
    
    On Error GoTo erro
    If Not MaskEdBox1.Text = "__:__" Then
        Call WritePrivateProfileString("When", "AlwaysAt", MaskEdBox1.Text, Arq)
        IniTime = TimeSerial(Hour(MaskEdBox1.Text), Minute(MaskEdBox1.Text), 0)
    Else
        Call WritePrivateProfileString("When", "AlwaysAt", "???", Arq)
        IniTime = vbEmpty
    End If
    
    If Not MaskEdBox2.Text = "__:__" Then
        Call WritePrivateProfileString("When", "Each", MaskEdBox2.Text, Arq)
        Interval = TimeSerial(Hour(MaskEdBox2.Text), Minute(MaskEdBox2.Text), 0)
    Else
        Call WritePrivateProfileString("When", "Each", "???", Arq)
        Interval = vbEmpty
    End If
    
    If OptionButton1 Then
        Call WritePrivateProfileString("When", "Default", False, Arq)
    Else
        Call WritePrivateProfileString("When", "Default", True, Arq)
    End If
    
    If OptionButton3 Then
        Call WritePrivateProfileString("When", "Days", "0", Arq)
    Else
        Days = 0
        Dim n As Byte
        For n = 0 To 6
            If CheckBox2(n + 1) Then Days = Days + 2 ^ n
        Next n
        Call WritePrivateProfileString("When", "Days", Days, Arq)
    End If
            
    If CheckBox1 Then
        Call WritePrivateProfileString("Log", "Save", "True", Arq)
    Else
        Call WritePrivateProfileString("Log", "Save", "False", Arq)
    End If
            
    If CheckBox4 Then
        Call WritePrivateProfileString("Backup", "Incremental", "True", Arq)
    Else
        Call WritePrivateProfileString("Backup", "Incremental", "False", Arq)
    End If
            
    Call WritePrivateProfileString("Destiny", "Dir", Text1.Text, Arq)
    
    For NLoops = 0 To List1.ListCount - 1
        If WritePrivateProfileString("Entries", CStr(NLoops), List1.List(NLoops), Arq) = 0 Then
            MsgBox "INI file full." & vbLf & "Last saved entry: " & List1.List(NLoops - 1), vbCritical
            GoTo Saída
        End If
    Next NLoops

    Screen.MousePointer = vbDefault
    
    Me.WindowState = vbMinimized

Saída:
    Exit Sub
    
erro:
    MsgBox Err.Number & vbLf & Err.Description, vbCritical
    Resume Saída

End Sub

Private Function VerificaDup(Item As String) As Boolean

    For LoopDup = 0 To List1.ListCount - 1
        If List1.List(LoopDup) = Item Then
            VerificaDup = True
            Exit Function
        End If
    Next LoopDup
    
    VerificaDup = False

End Function


Private Function VerifyErrors() As Boolean

    If List1.ListCount = 0 Then
        MsgBox "You must specify at least one file or directory for the backup!", vbCritical
        SSTab1.Tab = 1
        GoTo erro
    End If
    
    If Len(Text1.Text) = 0 Then
        MsgBox "You must specify the destiny dir.", vbCritical
        SSTab1.Tab = 2
        Text1.SetFocus
        GoTo erro
    ElseIf Text1.Text = "c:\" Or Text1.Text = "C:\" Then
        If MsgBox("The destiny dir was left as C:\." & vbLf & vbLf & "Confirm?", _
            vbYesNo + vbExclamation) = vbNo Then
            SSTab1.Tab = 2
            Text1.SetFocus
            GoTo erro
        End If
    ElseIf OptionButton1 And MaskEdBox1.Text = "__:__" Then
        MsgBox "You must specify a time for the backup!", vbCritical
        SSTab1.Tab = 0
        MaskEdBox1.SetFocus
        GoTo erro
    ElseIf OptionButton2 And MaskEdBox2.Text = "__:__" Then
        MsgBox "You must specify an interval for the backup!", vbCritical
        SSTab1.Tab = 0
        MaskEdBox2.SetFocus
        GoTo erro
    End If
    
    VerifyErrors = False

Saída:
    Exit Function
    
erro:
    VerifyErrors = True
    
End Function

Private Sub CheckBox2_Click(Index As Integer)

    OptionButton3.Value = False
    
End Sub


Private Sub Command1_Click()
On Error GoTo erro

    For NLoops = List1.ListCount - 1 To 0 Step -1
        If List1.Selected(NLoops) Then List1.RemoveItem (NLoops)
    Next NLoops
    
    
Saída:
    Exit Sub
    
erro:
    If Err.Number = 68 Then
        MsgBox "The selected drive is not available.", vbCritical
    Else
        MsgBox Err.Number & vbLf & Err.Description, vbCritical
    End If
    Resume Saída
    
End Sub


Private Sub Command2_Click()

    Unload Me

End Sub


Private Sub Command3_Click()

    If Not VerifyErrors Then SaveChanges
    
End Sub

Private Sub Command4_Click()

    If CheckBox3.Value = True Then
        Call AddItem(False, True)
    Else
        Call AddItem(False)
    End If

End Sub


Private Sub Command5_Click()

    AddItem (True)
    
End Sub

Private Sub Command6_Click()

    If MsgBox("This will effectuate the backup now!" & vbLf & vbLf & _
        "Confirm?", vbQuestion + vbYesNo) = vbYes Then Backup
                
End Sub



Private Sub Command7_Click()

    ShellExecute hwnd, "open", WindowsDir & "Log Autobak.txt", vbNullString, vbNullString, SW_SHOW

End Sub

Private Sub Dir2_Change()

    Text1.Text = Dir2.Path
    DestinyDir = Text1.Text

End Sub

Private Sub Drive1_Change()
On Error GoTo erro

    Dir1.Path = Drive1.Drive
        
Saída:
    Exit Sub
    
erro:
    If Err.Number = 68 Then
        MsgBox "The selected drive is not available.", vbCritical
        Drive1.Drive = "c:"
    Else
        MsgBox Err.Number & vbLf & Err.Description, vbCritical
    End If
    Resume Saída
    
End Sub

Private Sub Dir1_Change()
    
    File1.Path = Dir1.Path

End Sub


Private Sub Drive2_Change()
On Error GoTo erro
    
    Dir2.Path = Drive2.Drive
    
Saída:
    Exit Sub
    
erro:
    If Err.Number = 68 Then
        MsgBox "The selected drive is not available.", vbCritical
        Drive2.Drive = "c:"
    Else
        MsgBox Err.Number & vbLf & Err.Description, vbCritical
    End If
    Resume Saída

End Sub


Private Sub File1_DblClick()

    AddItem (True)
        
End Sub

Private Sub Form_Activate()

    If Not Default Then
        MaskEdBox1.SetFocus
    Else
        MaskEdBox2.SetFocus
    End If
    
    DoEvents
    
    If Not NoIniArchive Then Me.WindowState = vbMinimized
    
End Sub

Private Sub Form_Initialize()

    If App.PrevInstance Then
        MsgBox "There is another copy of the application being executed!", vbCritical
        OpenError = True
        Unload Me
        Set Form1 = Nothing
        End
    End If
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If ListWithFocus Then If KeyCode = 46 Then Command1_Click
        
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If

End Sub

Private Sub Form_Load()

    Dir1.Path = "C:\"
    Dir2.Path = "C:\"
    Initialize
    
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Auto Backup" & vbNullChar
    End With

    Shell_NotifyIcon NIM_ADD, nid
        
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Me.ScaleMode = vbPixels Then
        Msg = X
    Else
        Msg = X / Screen.TwipsPerPixelX
    End If

    Select Case Msg
        Case WM_LBUTTONUP '514 restore form window
        Me.WindowState = vbNormal
        Result = SetForegroundWindow(Me.hwnd)
        Me.Show
        Case WM_LBUTTONDBLCLK '515 restore form window
        Me.WindowState = vbNormal
        Result = SetForegroundWindow(Me.hwnd)
        Me.Show
        Case WM_RBUTTONUP '517 display popup menu
        Result = SetForegroundWindow(Me.hwnd)
        Me.PopupMenu Me.mnu_1
    End Select
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If OpenError Then Exit Sub
    
    If MsgBox("This will end the application." & vbLf & vbLf & "Are you sure?", vbQuestion + vbYesNo) = vbYes Then
        Unload Me
        Shell_NotifyIcon NIM_DELETE, nid
        Set Form1 = Nothing
        End
    Else
        Cancel = True
    End If
    
End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then Me.Hide
    
End Sub

Private Sub Label16_Click()

    ShellExecute hwnd, "open", "mailto:alb@cwb.matrix.com.br", vbNullString, vbNullString, SW_SHOW
    
End Sub

Private Sub List1_GotFocus()

    ListWithFocus = True

End Sub


Private Sub List1_LostFocus()

    ListWithFocus = False
    
End Sub


Private Sub MaskEdBox1_GotFocus()

    FieldFocus
    MskErr1 = False
    OptionButton1.Value = True
    
End Sub


Private Sub MaskEdBox1_LostFocus()
On Error GoTo erro

    If MskErr2 Or MaskEdBox1.Text = "__:__" Then Exit Sub
    
    IniTime = TimeSerial(Hour(MaskEdBox1.Text), Minute(MaskEdBox1.Text), 0)
            
Saída:
    Exit Sub
    
erro:
    If Err.Number = 13 Then
        MsgBox "Invalid time.", vbCritical
    Else
        MsgBox Err.Number & vbLf & Err.Description
    End If
    MskErr1 = True
    MaskEdBox1.SetFocus
    IniTime = vbEmpty
    Resume Saída
    
End Sub

Private Sub MaskEdBox2_GotFocus()

    OptionButton2.Value = True
    FieldFocus
    MskErr2 = False

End Sub

Sub FieldFocus()

    Screen.ActiveForm.ActiveControl.SelStart = 0
    Screen.ActiveForm.ActiveControl.SelLength = Len(Screen.ActiveForm.ActiveControl.Text)
    
End Sub
Private Sub MaskEdBox2_LostFocus()
On Error GoTo erro

    If MskErr1 Then Exit Sub
    
    If MaskEdBox2.Text = "__:__" Then
        OptionButton1.Value = True
        IniTime = "00:00"
        GoTo Saída
    End If

    Interval = TimeSerial(Hour(MaskEdBox2.Text), Minute(MaskEdBox2.Text), 0)
    
Saída:
    Exit Sub
    
erro:
    If Err.Number = 13 Then
        MsgBox "Invalid interval.", vbCritical
    Else
        MsgBox Err.Number & vbLf & Err.Description
    End If
    MskErr2 = True
    Interval = vbEmpty
    MaskEdBox2.SetFocus
    Resume Saída
    
End Sub





Private Sub MnuBackup_Click()

    Command6_Click
    
End Sub

Private Sub MnuRestaurar_Click()

    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
    
End Sub

Private Sub MnuSair_Click()

    Unload Me
    
End Sub

Private Sub MnuQuit_Click()

    Unload Me
    
End Sub

Private Sub MnuRestore_Click()

    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
    
End Sub

Private Sub OptionButton1_Click()

    MaskEdBox1.SetFocus
    
End Sub

Private Sub OptionButton2_Click()

    MaskEdBox2.SetFocus

End Sub

Private Sub OptionButton3_Click()

    For NLoops = 1 To 7
        CheckBox2(NLoops).Value = False
    Next NLoops
    
    OptionButton3.Value = True
    
End Sub

Private Sub Text1_GotFocus()

    FieldFocus
    
End Sub
Private Sub Timer1_Timer()

    If Interval = vbEmpty And IniTime = vbEmpty Then Exit Sub
        
    If Not OptionButton3 Then
        For NLoopsTimer = 1 To 7
            If CheckBox2(NLoopsTimer).Value = True Then If Format(Date, "w") = NLoopsTimer Then CheckTime
        Next NLoopsTimer
    Else
        CheckTime
    End If
        
End Sub
