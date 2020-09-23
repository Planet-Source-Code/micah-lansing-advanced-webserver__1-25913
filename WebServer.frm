VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form WebServer 
   Caption         =   "WebServer 1.2"
   ClientHeight    =   5295
   ClientLeft      =   3150
   ClientTop       =   1200
   ClientWidth     =   4860
   Icon            =   "WebServer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   4860
   Begin VB.Frame Frame4 
      Height          =   3675
      Left            =   60
      TabIndex        =   40
      Top             =   540
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton Command18 
         Caption         =   "X"
         Height          =   195
         Left            =   4440
         TabIndex        =   63
         Top             =   60
         Width           =   195
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00800000&
         FillColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   4635
         TabIndex        =   69
         Top             =   0
         Width           =   4695
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3315
         Left            =   60
         TabIndex        =   41
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   5847
         _Version        =   327681
         TabHeight       =   520
         TabCaption(0)   =   "IPs"
         TabPicture(0)   =   "WebServer.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label9"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label10"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "List2"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "List3"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Message"
         TabPicture(1)   =   "WebServer.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label11"
         Tab(1).Control(1)=   "Label15"
         Tab(1).Control(2)=   "List4"
         Tab(1).Control(3)=   "MessageTxt"
         Tab(1).Control(4)=   "Command7"
         Tab(1).Control(5)=   "Command8"
         Tab(1).ControlCount=   6
         TabCaption(2)   =   "Misc."
         TabPicture(2)   =   "WebServer.frx":0342
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label12"
         Tab(2).Control(1)=   "Label13"
         Tab(2).Control(2)=   "Label14"
         Tab(2).Control(3)=   "Shape2"
         Tab(2).Control(4)=   "Label16"
         Tab(2).Control(5)=   "Label17"
         Tab(2).Control(6)=   "MaxSock"
         Tab(2).Control(7)=   "Command10"
         Tab(2).Control(8)=   "Command11"
         Tab(2).Control(9)=   "Command12"
         Tab(2).Control(10)=   "Command13"
         Tab(2).Control(11)=   "SetSocks"
         Tab(2).Control(12)=   "Command15"
         Tab(2).Control(13)=   "Command17"
         Tab(2).Control(14)=   "Timeoutcmd"
         Tab(2).Control(15)=   "Command19"
         Tab(2).Control(16)=   "Command20"
         Tab(2).ControlCount=   17
         Begin VB.CommandButton Command20 
            Caption         =   "?"
            Height          =   255
            Left            =   -71940
            TabIndex        =   65
            Top             =   1140
            Width           =   375
         End
         Begin VB.CommandButton Command19 
            Caption         =   "?"
            Height          =   255
            Left            =   -71940
            TabIndex        =   64
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton Timeoutcmd 
            Caption         =   "Off"
            Height          =   255
            Left            =   -72660
            TabIndex        =   62
            Top             =   1140
            Width           =   735
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Off"
            Height          =   255
            Left            =   -72660
            TabIndex        =   61
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Infinite"
            Height          =   315
            Left            =   -74100
            TabIndex        =   60
            Top             =   2820
            Width           =   675
         End
         Begin VB.CommandButton SetSocks 
            Caption         =   "Set"
            Height          =   315
            Left            =   -74760
            TabIndex        =   59
            Top             =   2820
            Width           =   675
         End
         Begin VB.CommandButton Command13 
            Caption         =   ">"
            Height          =   255
            Left            =   -73740
            TabIndex        =   58
            Top             =   2460
            Width           =   255
         End
         Begin VB.CommandButton Command12 
            Caption         =   ">>"
            Height          =   255
            Left            =   -74100
            TabIndex        =   57
            Top             =   2460
            Width           =   375
         End
         Begin VB.CommandButton Command11 
            Caption         =   "<<"
            Height          =   255
            Left            =   -74460
            TabIndex        =   56
            Top             =   2460
            Width           =   375
         End
         Begin VB.CommandButton Command10 
            Caption         =   "<"
            Height          =   255
            Left            =   -74700
            TabIndex        =   55
            Top             =   2460
            Width           =   255
         End
         Begin VB.TextBox MaxSock 
            Height          =   285
            Left            =   -74400
            Locked          =   -1  'True
            TabIndex        =   54
            Text            =   "Infinite"
            Top             =   2100
            Width           =   615
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   -72300
            TabIndex        =   50
            Top             =   2820
            Width           =   915
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Send"
            Height          =   375
            Left            =   -74040
            TabIndex        =   49
            Top             =   2820
            Width           =   915
         End
         Begin VB.TextBox MessageTxt 
            Height          =   1635
            Left            =   -72960
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   47
            Top             =   960
            Width           =   2415
         End
         Begin VB.ListBox List4 
            Height          =   1815
            ItemData        =   "WebServer.frx":035E
            Left            =   -74820
            List            =   "WebServer.frx":0360
            TabIndex        =   46
            Top             =   780
            Width           =   1815
         End
         Begin VB.ListBox List3 
            Height          =   1815
            ItemData        =   "WebServer.frx":0362
            Left            =   2580
            List            =   "WebServer.frx":0364
            TabIndex        =   43
            Top             =   780
            Width           =   1815
         End
         Begin VB.ListBox List2 
            Height          =   1815
            ItemData        =   "WebServer.frx":0366
            Left            =   180
            List            =   "WebServer.frx":0368
            TabIndex        =   42
            Top             =   780
            Width           =   1815
         End
         Begin VB.Label Label17 
            Caption         =   "Too many loaded winsock controls can casue ""Out of memory"" error and or system halt."
            Height          =   615
            Left            =   -73080
            TabIndex        =   68
            Top             =   2340
            Width           =   2355
         End
         Begin VB.Label Label16 
            Caption         =   "Warning:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   -72840
            MousePointer    =   2  'Cross
            TabIndex        =   67
            Top             =   1980
            Width           =   1215
         End
         Begin VB.Label Label15 
            Height          =   195
            Left            =   -72900
            TabIndex        =   66
            Top             =   480
            Width           =   975
         End
         Begin VB.Shape Shape2 
            Height          =   1155
            Left            =   -74820
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label Label14 
            Caption         =   "Winsock Controls:"
            Height          =   195
            Left            =   -74820
            TabIndex        =   53
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label13 
            Caption         =   "Connection Timeout"
            Height          =   255
            Left            =   -74880
            TabIndex        =   52
            Top             =   1200
            Width           =   1515
         End
         Begin VB.Label Label12 
            Caption         =   "/.. Protection"
            Height          =   255
            Left            =   -74820
            TabIndex        =   51
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "0.0.0.0"
            Height          =   195
            Left            =   -72900
            TabIndex        =   48
            Top             =   720
            Width           =   1995
         End
         Begin VB.Label Label10 
            Caption         =   "Blocked"
            Height          =   195
            Left            =   3120
            TabIndex        =   45
            Top             =   540
            Width           =   795
         End
         Begin VB.Label Label9 
            Caption         =   "Unblocked"
            Height          =   195
            Left            =   540
            TabIndex        =   44
            Top             =   540
            Width           =   1035
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2910
      Left            =   1200
      TabIndex        =   29
      Top             =   300
      Visible         =   0   'False
      Width           =   2415
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00800000&
         FillColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   2355
         TabIndex        =   70
         Top             =   0
         Width           =   2415
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Save IPs"
         Height          =   375
         Left            =   1455
         TabIndex        =   35
         Top             =   2460
         Width           =   795
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Ok"
         Height          =   375
         Left            =   135
         TabIndex        =   34
         Top             =   2460
         Width           =   795
      End
      Begin VB.TextBox Text2 
         Height          =   1035
         Left            =   75
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   1350
         Width           =   2235
      End
      Begin VB.Label Label8 
         Caption         =   "0"
         Height          =   180
         Left            =   1140
         TabIndex        =   39
         Top             =   900
         Width           =   1200
      End
      Begin VB.Label Label7 
         Caption         =   "Errors:"
         Height          =   210
         Left            =   75
         TabIndex        =   38
         Top             =   900
         Width           =   1050
      End
      Begin VB.Label PagesViewed 
         Caption         =   "0"
         Height          =   195
         Left            =   1140
         TabIndex        =   37
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label BytesSentLb 
         Caption         =   "0"
         Height          =   195
         Left            =   1140
         TabIndex        =   36
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "IPs:"
         Height          =   255
         Left            =   90
         TabIndex        =   32
         Top             =   1125
         Width           =   435
      End
      Begin VB.Label Label5 
         Caption         =   "Pages Viewed:"
         Height          =   195
         Left            =   75
         TabIndex        =   31
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Bytes Sent:"
         Height          =   255
         Left            =   75
         TabIndex        =   30
         Top             =   420
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4635
      Left            =   300
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00800000&
         FillColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   4455
         TabIndex        =   71
         Top             =   60
         Width           =   4515
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Load"
         Height          =   255
         Left            =   2520
         TabIndex        =   27
         Top             =   2940
         Width           =   735
      End
      Begin VB.CommandButton Ok 
         Caption         =   "OK"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2520
         TabIndex        =   26
         Top             =   3720
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Default"
         Height          =   375
         Left            =   3120
         TabIndex        =   25
         Top             =   4140
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1680
         TabIndex        =   24
         Top             =   4140
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Set Errors"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   4140
         Width           =   975
      End
      Begin VB.ListBox List1 
         Height          =   1035
         ItemData        =   "WebServer.frx":036A
         Left            =   120
         List            =   "WebServer.frx":036C
         TabIndex        =   22
         Top             =   2940
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   2475
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   21
         Top             =   360
         Width           =   4095
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   4320
         Y1              =   4020
         Y2              =   4020
      End
   End
   Begin VB.CommandButton Hidecmd 
      Caption         =   "Hide Server"
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   3180
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Path..."
      Height          =   3015
      Left            =   480
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   1920
         TabIndex        =   18
         Top             =   240
         Width           =   1815
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1815
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton CancelPath 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton Okcmd 
         Caption         =   "OK"
         Height          =   375
         Left            =   720
         TabIndex        =   14
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox PathText 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Text            =   "C:\Webserver\index.html"
         Top             =   2160
         Width           =   3495
      End
   End
   Begin MSWinsockLib.Winsock HTTP 
      Index           =   0
      Left            =   360
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
      LocalPort       =   80
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   60000
      Left            =   4200
      Top             =   4440
   End
   Begin VB.CommandButton Dissable 
      Caption         =   "Enabled"
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox cnt2 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0"
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox cnt1 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0"
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton IPScmd 
      Caption         =   "IPs"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   3180
      Width           =   615
   End
   Begin VB.CommandButton Disconnect 
      Caption         =   "Stop"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton Reset 
      Caption         =   "Restet"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1380
      TabIndex        =   3
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton connect 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox IPtxt 
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox HTML 
      Height          =   2895
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   -60
      TabIndex        =   28
      Top             =   2940
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label3 
      Caption         =   "Connection Timeout Disabled."
      Height          =   375
      Left            =   3420
      TabIndex        =   10
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   3240
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Current Connections"
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Connections made"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Connect1 
         Caption         =   "Start"
      End
      Begin VB.Menu Disconnect1 
         Caption         =   "Stop"
      End
      Begin VB.Menu MyIP 
         Caption         =   "My IP"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Code 
      Caption         =   "Code"
      Begin VB.Menu PathMenu 
         Caption         =   "Path"
      End
      Begin VB.Menu Errors404 
         Caption         =   "HTTP 404 Errors"
      End
      Begin VB.Menu Stats 
         Caption         =   "Stats"
      End
   End
   Begin VB.Menu Security 
      Caption         =   "Security"
      Begin VB.Menu Tabs 
         Caption         =   "Tabs"
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
      Begin VB.Menu Help 
         Caption         =   "Help"
      End
      Begin VB.Menu Author 
         Caption         =   "Author"
      End
   End
   Begin VB.Menu TrayMenu 
      Caption         =   "TrayMenu"
      Visible         =   0   'False
      Begin VB.Menu Hidemnu 
         Caption         =   "Hide"
      End
      Begin VB.Menu Showmnu 
         Caption         =   "Show"
         Enabled         =   0   'False
      End
      Begin VB.Menu Startmnu 
         Caption         =   "Start"
      End
      Begin VB.Menu Stopmnu 
         Caption         =   "Stop"
         Enabled         =   0   'False
      End
      Begin VB.Menu Exitmnu 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "WebServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As String
Dim AllErrors As String
Dim AllIPs As String
Dim B As String
Dim Blocked_IPs As String
Dim C As String
Dim D As Integer
Dim CodeType As String
Dim Connected As String
Dim ConnectCnt As Integer
Dim ConnectionCounter As Integer
Dim Counter As Integer
Dim Data As String
Dim Error As String
Dim Error1, Error2, Error3, Error4, Error5, Error6, Error1A, Error2A, Error3A, Error4A, Error5A, Error6A As String
Dim Errorcnt As Integer
Dim Final As String
Dim HitCount As String
Dim HtmlIndex As String
Dim HtmlIndexFile As String
Dim i%
Dim IPs As String
Dim IPList As String
Dim Leftty As String
Dim MsgIP As String
Dim Message, msg, Buttons, Title As String
Dim MyPic As String
Dim NewData As String
Dim Oldheight, Diffrence, Newheight As String
Dim OldIPs As String
Dim Oldwidth, WNewwidth, WDiffrence As String
Dim OneIP As String
Dim Request As String
Dim Request1 As String
Dim RequestError As String
Dim Righty As String
Dim Path As String
Dim s%
Dim SlashDot As Integer
Dim SelectedItem As String
Dim SendMsg As Integer
Dim Setup As String
Dim SetInterval As Integer
Dim Socks As String
Dim Stuff As String
Dim Stufftwo As String
Dim ThisMessage As String
Dim TrayMsg As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type mouseptr
  pos As POINTAPI
  xStart As Long
  yStart As Long
  width As Long
  height As Long
  moving As Boolean
  Left As Long
  Top As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Dim mouse As mouseptr

Option Explicit

'API constants
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
Private Const WM_LBUTTONDBLCLK = &H203   'Double click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up
Private Const WM_RBUTTONDBLCLK = &H206   'Double click

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As Stuff) As Boolean

Private Tray As Stuff
Private Type Stuff
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    MenuMsg As Long
    TIcon As Long
    msgTip As String * 64
End Type


Private Sub Author_Click()
MsgBox "Written by Micah Lansing, © 2001", vbOKOnly, "Webserver 1.2"
End Sub

Private Sub BytesSentLb_Click()
Frame3.ZOrder 0
End Sub

Private Sub CancelPath_Click()
connect.Enabled = True
Connect1.Enabled = True
Frame1.Visible = False
End Sub

Private Sub Command1_Click()
If D = 1 Then
msg = "Save HTTP 404 Errors?"
Buttons = vbYesNoCancel
Title = "Save?"
Message = MsgBox(msg, Buttons, Title)
    If Message = vbYes Then
    AllErrors = Error1A & " || " & Error2A & " || " & Error3A & " || " & Error4A & " || " & Error5A & " || " & Error6A
        On Error Resume Next
        Kill Path + "\Errors.err" 'deletes old custiom errors
        Open Path + "\Errors.err" For Binary Access Write As #1 'Saves custom errors to a file
            Put 1, , AllErrors
        Close #1
        Error1 = Error1A
Error2 = Error2A
Error3 = Error3A
Error4 = Error4A
Error5 = Error5A
Error6 = Error6A
    End If
    If Message = vbNo Then
    Error1 = Error1A
Error2 = Error2A
Error3 = Error3A
Error4 = Error4A
Error5 = Error5A
Error6 = Error6A
    End If
    If Message = vbCancel Then GoTo 12
End If
Error1 = Error1A
Error2 = Error2A
Error3 = Error3A
Error4 = Error4A
Error5 = Error5A
Error6 = Error6A
12
Frame2.Visible = False
Command1.Enabled = False
End Sub

Private Sub Command10_Click()
    If MaxSock.Text = "Infinite" Then MaxSock.Text = 5
MaxSock.Text = MaxSock.Text - 1
If MaxSock.Text < 5 Then MaxSock.Text = 5
End Sub

Private Sub Command11_Click()
    If MaxSock.Text = "Infinite" Then MaxSock.Text = 5
MaxSock.Text = MaxSock.Text - 10
If MaxSock.Text < 5 Then MaxSock.Text = 5
End Sub

Private Sub Command12_Click()
    If MaxSock.Text = "Infinite" Then MaxSock.Text = 5
MaxSock.Text = MaxSock.Text + 10
End Sub

Private Sub Command13_Click()
    If MaxSock.Text = "Infinite" Then MaxSock.Text = 5
MaxSock.Text = MaxSock.Text + 1
End Sub

Private Sub Command15_Click()
MaxSock.Text = "Infinite"
Frame4.ZOrder 0
End Sub

Private Sub Command17_Click()
If SlashDot = 0 Then
    SlashDot = 1
    Command17.Caption = "On"
Else
    SlashDot = 0
    Command17.Caption = "Off"
End If
End Sub

Private Sub Command18_Click()
Frame4.Visible = False
End Sub

Private Sub Command19_Click()
MsgBox "/.. Protection stops the client from using /.. to go back a directory on your hard drive.", vbOKOnly, "/.. Protection"
End Sub

Private Sub Command2_Click()
Frame2.Visible = False
End Sub

Private Sub Command20_Click()
MsgBox "Turns on or off a one minute connection timer. If on, connections are only held for one minute before being closed.", vbOKOnly, "Connection Timeout"
End Sub

Private Sub Command3_Click()
DftError
Frame2.Visible = False
End Sub

Private Sub Command4_Click()
Frame3.Visible = False
End Sub

Private Sub Command5_Click()
Open Path & "\errors.err" For Binary Access Read As #1 'Loads saved custiom errors
    AllErrors = Space(LOF(1))
    Get 1, , AllErrors
Close #1
If AllErrors = "" Then Kill Path + "/errors.err": Exit Sub
Error1A = Mid(AllErrors, 1, InStr(1, AllErrors, " || ") - 1)
AllErrors = Mid(AllErrors, InStr(1, AllErrors, " || ") + 4)
Error2A = Mid(AllErrors, 1, InStr(1, AllErrors, " || ") - 1)
AllErrors = Mid(AllErrors, InStr(1, AllErrors, " || ") + 4)
Error3A = Mid(AllErrors, 1, InStr(1, AllErrors, " || ") - 1)
AllErrors = Mid(AllErrors, InStr(1, AllErrors, " || ") + 4)
Error4A = Mid(AllErrors, 1, InStr(1, AllErrors, " || ") - 1)
AllErrors = Mid(AllErrors, InStr(1, AllErrors, " || ") + 4)
Error5A = Mid(AllErrors, 1, InStr(1, AllErrors, " || ") - 1)
AllErrors = Mid(AllErrors, InStr(1, AllErrors, " || ") + 4)
Error6A = AllErrors
AllErrors = ""
Command1.Enabled = True
D = 0
Frame2.ZOrder 0
End Sub

Private Sub Command6_Click()
Open Path + "/ServerIpLog.txt" For Binary Access Read Write As #1 'saves IPs to file
  OldIPs = Space(LOF(1))
  Get 1, , OldIPs
  OldIPs = OldIPs & Chr(13) & Chr(10) & "---------------" & Chr(13) & Chr(10) & Text2.Text
  Put 1, , OldIPs
Close #1
Frame4.ZOrder 0
End Sub

Private Sub Command7_Click()
Label15.Caption = "Not yet sent"
SendMsg = 1
End Sub

Private Sub Command8_Click()
Label15.Caption = ""
Label11.Caption = "0.0.0.0"
MessageTxt.Text = ""
SendMsg = 0
End Sub

Private Sub Dir1_Click()
Frame1.ZOrder 0
End Sub

Private Sub Errors404_Click()
Frame2.Visible = True
Frame2.ZOrder (0) 'sets frame2 on top
Frame2.Refresh
Picture3.CurrentX = 0
Picture3.CurrentY = 0
Picture3.Print "HTTP 404 Errors"
End Sub

Private Sub Exitmnu_Click()
Unload Me
End Sub

Private Sub File1_Click()
Frame1.ZOrder 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'set up the TrayMsg
    If Me.ScaleMode = vbPixels Then
        TrayMsg = X
    Else
        TrayMsg = X / Screen.TwipsPerPixelX
    End If

    Select Case TrayMsg

        Case WM_RBUTTONUP 'right click
            Me.PopupMenu TrayMenu
    
        Case WM_LBUTTONDBLCLK  'when the left mouse button is dubble clicked
            Me.Show
            Me.SetFocus
            Showmnu.Enabled = False
            Hidemnu.Enabled = True
    End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next
If WebServer.height < 5985 Then
    WebServer.height = 5985
    Oldheight = 5985
    Newheight = Oldheight
    If WebServer.width < 4980 Then
        WebServer.width = 4980
        Oldwidth = 4980
        WNewwidth = Oldwidth: GoTo 15
    End If
GoTo 14
End If
If WebServer.width < 4980 Then
        WebServer.width = 4980
        Oldwidth = 4980
        WNewwidth = Oldwidth: GoTo 15
End If
Newheight = WebServer.height
Diffrence = Newheight - Oldheight
If Diffrence = 0 Then GoTo 14
connect.Top = connect.Top + Diffrence
Reset.Top = Reset.Top + Diffrence
Disconnect.Top = Disconnect.Top + Diffrence
Dissable.Top = Dissable.Top + Diffrence
Shape1.Top = Shape1.Top + Diffrence
IPScmd.Top = IPScmd.Top + Diffrence
Hidecmd.Top = Hidecmd.Top + Diffrence
cnt1.Top = cnt1.Top + Diffrence
cnt2.Top = cnt2.Top + Diffrence
Label1.Top = Label1.Top + Diffrence
Label2.Top = Label2.Top + Diffrence
Label3.Top = Label3.Top + Diffrence
IPtxt.Top = IPtxt.Top + Diffrence
ProgressBar1.Top = ProgressBar1.Top + Diffrence
Frame1.Top = Frame1.Top + (1 / 2 * Diffrence)
Frame2.Top = Frame2.Top + (1 / 2 * Diffrence)
Frame3.Top = Frame3.Top + (1 / 2 * Diffrence)
Frame4.Top = Frame4.Top + (1 / 2 * Diffrence)
HTML.height = HTML.height + Diffrence

14
WNewwidth = WebServer.width
WDiffrence = WNewwidth - Oldwidth
If WDiffrence = 0 Then GoTo 15
If connect.Left < 0 Then GoTo 13
connect.Left = connect.Left + ((1 / 2 * WDiffrence) - Diffrence)
Reset.Left = Reset.Left + ((1 / 2 * WDiffrence) - Diffrence)
Disconnect.Left = Disconnect.Left + ((1 / 2 * WDiffrence) - Diffrence)
13
Dissable.Left = Dissable.Left + WDiffrence
Shape1.Left = Shape1.Left + WDiffrence
IPScmd.Left = IPScmd.Left + WDiffrence
Hidecmd.Left = Hidecmd.Left + WDiffrence
cnt1.Left = cnt1.Left + WDiffrence
cnt2.Left = cnt2.Left + WDiffrence
Label1.Left = Label1.Left + WDiffrence
Label2.Left = Label2.Left + WDiffrence
Label3.Left = Label3.Left + WDiffrence
ProgressBar1.width = ProgressBar1.width + WDiffrence
IPtxt.width = IPtxt.width + WDiffrence
HTML.width = HTML.width + WDiffrence
Frame1.Left = Frame1.Left + (1 / 2 * WDiffrence)
Frame2.Left = Frame2.Left + (1 / 2 * WDiffrence)
Frame3.Left = Frame3.Left + (1 / 2 * WDiffrence)
Frame4.Left = Frame4.Left + (1 / 2 * WDiffrence)
15
Oldwidth = WNewwidth
Oldheight = Newheight
If Shape1.Left < 3240 Then Shape1.Left = 3240
If Shape1.Left - 840 < 2499 Then
    connect.Left = 120
    Disconnect.Left = 2400
    Reset.Left = 1320
End If
If connect.Left < 120 Then connect.Left = 120
If Disconnect.Left < 2400 Then Disconnect.Left = 2400
If Reset.Left < 1320 Then Reset.Left = 1320
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, Tray
On Error Resume Next
Unload HelpForm
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mouse.moving = True
  Call GetCursorPos(mouse.pos)
  mouse.xStart = mouse.pos.X
  mouse.yStart = mouse.pos.Y
  mouse.Left = Frame1.Left
  mouse.Top = Frame1.Top
  Frame1.ZOrder 0
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If mouse.moving = True Then
    Call GetCursorPos(mouse.pos)
    If Abs(mouse.xStart - mouse.pos.X) > 5 Or Abs(mouse.yStart - mouse.pos.Y) > 5 Then
      Frame1.Move mouse.Left + (mouse.pos.X - mouse.xStart) * Screen.TwipsPerPixelX, mouse.Top + (mouse.pos.Y - mouse.yStart) * Screen.TwipsPerPixelY
    Else
      Frame1.Move mouse.Left, mouse.Top
    End If
  End If
End Sub

Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mouse.moving = False
  Picture1.CurrentX = 0
Picture1.CurrentY = 0
Picture1.Print "Security"
Picture2.CurrentX = 0
Picture2.CurrentY = 0
Picture2.Print "Stats"
Picture3.CurrentX = 0
Picture3.CurrentY = 0
Picture3.Print "HTTP 404 Errors"
End Sub


Private Sub Frame2_Click()
Frame2.ZOrder 0
End Sub

Private Sub Frame3_Click()
Frame3.ZOrder 0
End Sub

Private Sub Help_Click()
HelpForm.Show (vbModeless) 'shows the helpform
End Sub

Private Sub Hidecmd_Click()
Me.Hide 'hides this form
Hidemnu.Enabled = False
Showmnu.Enabled = True
End Sub

Private Sub Hidemnu_Click()
Me.Hide
Hidemnu.Enabled = False
Showmnu.Enabled = True
End Sub

Private Sub HTTP_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'If there is an error, close or unload the winsock
If Index <> 0 And Socks = "I" Then
    Unload HTTP(Index)
    Unload Timer1(Index)
Else
    HTTP(Index).Close
End If
End Sub

Private Sub HTTP_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
ProgressBar1.Max = bytesRemaining + bytesSent
ProgressBar1.Value = bytesSent
If bytesRemaining = 0 Then ProgressBar1.Value = 0
End Sub

Private Sub Label12_Click()
Frame4.ZOrder 0
End Sub

Private Sub Label13_Click()
Frame4.ZOrder 0
End Sub

Private Sub Label14_Click()
Frame4.ZOrder 0
End Sub

Private Sub Label16_Click()
MsgBox "Note: Too many loaded winsock controls can casue 'Out of memory' error and or system halt. Although the more Winsocks you have loaded, the better the server will perform.", vbOKOnly, "Winsock"
Frame4.ZOrder 0
End Sub

Private Sub Label17_Click()
Frame4.ZOrder 0
End Sub

Private Sub Label4_Click()
Frame3.ZOrder 0
End Sub

Private Sub Label5_Click()
Frame3.ZOrder 0
End Sub

Private Sub Label6_Click()
Frame3.ZOrder 0
End Sub

Private Sub Label7_Click()
Frame3.ZOrder 0
End Sub

Private Sub Label8_Click()
Frame3.ZOrder 0
End Sub

Private Sub List1_DblClick()
SelectedItem = List1.List(List1.ListIndex)
Select Case SelectedItem
    Case "Error 1"
        Text1.Text = Error1A
    Case "Error 2"
        Text1.Text = Error2A
    Case "Error 3"
        Text1.Text = Error3A
    Case "Error 4"
        Text1.Text = Error4A
    Case "Error 5"
        Text1.Text = Error5A
    Case "Error 6"
        Text1.Text = Error6A
End Select
Ok.Enabled = True
Text1.Enabled = True
Text1.SetFocus
List1.Enabled = False
Frame2.ZOrder 0
End Sub

Private Sub List2_DblClick()
List3.AddItem List2.List(List2.ListIndex)
Blocked_IPs = Blocked_IPs & List2.List(List2.ListIndex) & " " 'Add the IP to the block list
List2.RemoveItem (List2.ListIndex)
Frame4.ZOrder 0
End Sub

Private Sub List3_DblClick()
Dim OneBlocked As String
Dim AllBlocked As String
Dim SomeBlocked As String
List2.AddItem List3.List(List3.ListIndex)
AllBlocked = Blocked_IPs

Do 'Adds IPs to list, and makes sure there Isnt any Dubbles
OneBlocked = Mid(AllBlocked, 1, InStr(1, AllBlocked, " ") - 1)
AllBlocked = Mid(AllBlocked, InStr(1, AllBlocked, " ") + 1)

If OneBlocked <> List3.List(List3.ListIndex) Then 'Check to see if the IP matches
    SomeBlocked = SomeBlocked & OneBlocked & " "
Else
GoTo 25
End If
Loop Until AllBlocked = ""
25
Blocked_IPs = SomeBlocked & AllBlocked 'put the new Block list togather
List3.RemoveItem (List3.ListIndex)
Frame4.ZOrder 0
End Sub

Private Sub List4_DblClick()
Label11.Caption = List4.List(List4.ListIndex)
Frame4.ZOrder 0
End Sub

Private Sub MessageTxt_Click()
Frame4.ZOrder 0
End Sub

Private Sub MyIP_Click()
MsgIP = HTTP(0).LocalIP 'gets you computers IP address
ThisMessage = MsgBox(MsgIP, vbOKOnly, "Your IP is...")
End Sub

Private Sub Ok_Click()
SelectedItem = List1.List(List1.ListIndex)
Select Case SelectedItem
    Case "Error 1"
        Error1A = Text1.Text
    Case "Error 2"
        Error2A = Text1.Text
    Case "Error 3"
        Error3A = Text1.Text
    Case "Error 4"
        Error4A = Text1.Text
    Case "Error 5"
        Error5A = Text1.Text
    Case "Error 6"
        Error6A = Text1.Text
End Select
Command1.Enabled = True
Ok.Enabled = False
Text1.Enabled = False
List1.Enabled = True
D = 1
Frame2.ZOrder 0
End Sub


Private Sub PagesViewed_Click()
Frame3.ZOrder 0
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mouse.moving = True
  Call GetCursorPos(mouse.pos)
  mouse.xStart = mouse.pos.X
  mouse.yStart = mouse.pos.Y
  mouse.Left = Frame4.Left
  mouse.Top = Frame4.Top
  Frame4.ZOrder 0
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If mouse.moving = True Then
    Call GetCursorPos(mouse.pos)
    If Abs(mouse.xStart - mouse.pos.X) > 5 Or Abs(mouse.yStart - mouse.pos.Y) > 5 Then
      Frame4.Move mouse.Left + (mouse.pos.X - mouse.xStart) * Screen.TwipsPerPixelX, mouse.Top + (mouse.pos.Y - mouse.yStart) * Screen.TwipsPerPixelY
    Else
      Frame4.Move mouse.Left, mouse.Top
    End If
  End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mouse.moving = False
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mouse.moving = True
  Call GetCursorPos(mouse.pos)
  mouse.xStart = mouse.pos.X
  mouse.yStart = mouse.pos.Y
  mouse.Left = Frame3.Left
  mouse.Top = Frame3.Top
  Frame3.ZOrder 0
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If mouse.moving = True Then
    Call GetCursorPos(mouse.pos)
    If Abs(mouse.xStart - mouse.pos.X) > 5 Or Abs(mouse.yStart - mouse.pos.Y) > 5 Then
      Frame3.Move mouse.Left + (mouse.pos.X - mouse.xStart) * Screen.TwipsPerPixelX, mouse.Top + (mouse.pos.Y - mouse.yStart) * Screen.TwipsPerPixelY
    Else
      Frame3.Move mouse.Left, mouse.Top
    End If
  End If
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mouse.moving = False
End Sub
Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mouse.moving = True
  Call GetCursorPos(mouse.pos)
  mouse.xStart = mouse.pos.X
  mouse.yStart = mouse.pos.Y
  mouse.Left = Frame2.Left
  mouse.Top = Frame2.Top
  Frame2.ZOrder 0
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If mouse.moving = True Then
    Call GetCursorPos(mouse.pos)
    If Abs(mouse.xStart - mouse.pos.X) > 5 Or Abs(mouse.yStart - mouse.pos.Y) > 5 Then
      Frame2.Move mouse.Left + (mouse.pos.X - mouse.xStart) * Screen.TwipsPerPixelX, mouse.Top + (mouse.pos.Y - mouse.yStart) * Screen.TwipsPerPixelY
    Else
      Frame2.Move mouse.Left, mouse.Top
    End If
  End If
End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mouse.moving = False
End Sub

Private Sub reset_Click()
'Reset the server
HTTP(0).Close
    Open Path + HtmlIndexFile For Binary Access Read As #1 'Opens index file
        Data = Space(LOF(1))
        Get #1, , Data
        HTML.Text = Data 'sets file data to textbox
        Close #1
    NewData = ""
    IPtxt = ""
    HTML.Text = Data
On Error Resume Next
If Socks = "I" Then
    For i% = 1 To ConnectCnt 'Unloads winsocks and timers
      Unload HTTP(i%)
      Unload Timer1(i%)
    Next i%
End If
    ConnectCnt = 0
    Counter = 0
    ConnectionCounter = 0
    cnt1 = Counter
    cnt2 = ConnectionCounter
    HTTP(0).Listen
End Sub
Private Sub Connect_Click()
'Starts the server
HTTP(0).Listen
    Connected = "YES"
    Connect1.Enabled = False
    connect.Enabled = False
    Reset.Enabled = True
    Disconnect.Enabled = True
    Disconnect1.Enabled = True
Command10.Enabled = False
Command11.Enabled = False
Command12.Enabled = False
Command13.Enabled = False
Startmnu.Enabled = False
Stopmnu.Enabled = True
SetSocks.Enabled = False
Command15.Enabled = False
PathMenu.Enabled = False
End Sub

Private Sub Connect1_Click()
'Starts the server
HTTP(0).Listen
    Connected = "YES"
    Connect1.Enabled = False
    connect.Enabled = False
    Reset.Enabled = True
    Disconnect.Enabled = True
    Disconnect1.Enabled = True
Command10.Enabled = False
Command11.Enabled = False
Command12.Enabled = False
Command13.Enabled = False
Startmnu.Enabled = False
Stopmnu.Enabled = True
SetSocks.Enabled = False
Command15.Enabled = False
PathMenu.Enabled = False
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1
PathText = Dir1
End Sub
Private Sub Disconnect_Click()
HTTP(0).Close
    Open Path + HtmlIndexFile For Binary Access Read As #1 'Opens index file
        Data = Space(LOF(1))
        Get #1, , Data
        HTML.Text = Data 'Sets file data to textbox
        Close #1
    NewData = ""
    IPtxt = ""
    HTML.Text = Data
    connect.Enabled = True
    Connect1.Enabled = True
    Reset.Enabled = False
    Disconnect.Enabled = False
    Disconnect1.Enabled = False
    Stopmnu.Enabled = False
    Startmnu.Enabled = True
    Command10.Enabled = True
    Command11.Enabled = True
    Command12.Enabled = True
    Command13.Enabled = True
    SetSocks.Enabled = True
    Command15.Enabled = True
On Error Resume Next
If Socks = "I" Then
    For i% = 1 To ConnectCnt 'Unload winsocks and timers
      HTTP(i%).Close
      Unload HTTP(i%)
      Unload Timer1(i%)
    Next i%
End If
    ConnectCnt = 0
    Counter = 0
    ConnectionCounter = 0
    cnt1 = Counter
    cnt2 = ConnectionCounter
PathMenu.Enabled = True
End Sub

Private Sub Disconnect1_Click()
HTTP(0).Close
    Open Path + HtmlIndexFile For Binary Access Read As #1 'opens index file
        Data = Space(LOF(1))
        Get #1, , Data
        HTML.Text = Data 'sets data of index file to textbox
        Close #1
    NewData = ""
    IPtxt = ""
    HTML.Text = Data
    connect.Enabled = True
    Connect1.Enabled = True
    Startmnu.Enabled = True
    Stopmnu.Enabled = False
    Reset.Enabled = False
    Disconnect.Enabled = False
    Disconnect1.Enabled = False
    On Error Resume Next
If Socks = "I" Then
    For i% = 1 To ConnectCnt 'Unload winsocks and timers
        HTTP(i%).Close
        Unload HTTP(i%)
        Unload Timer1(i%)
    Next i%
End If
    ConnectCnt = 0
    Counter = 0
    ConnectionCounter = 0
    cnt1 = Counter
    cnt2 = ConnectionCounter
    PathMenu.Enabled = True
End Sub

Private Sub Dissable_Click()
If B = 0 Then 'Turns off the connection timeout
B = 1
C = 1
Dissable.Caption = "Enabled"
Timeoutcmd.Caption = "Off"
Label3.Caption = "Connection Timeout Disabled."
Timer1(0).Enabled = False
Else 'turns on the connection timeout
Dissable.Caption = "Disable"
Label3.Caption = "Connection Timeout Enabled."
Timeoutcmd.Caption = "On"
Timer1(0).Enabled = True
C = 0
B = 0
End If
End Sub
Private Sub Drive1_Change()
Dir1.Path = Drive1
End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Private Sub File1_DblClick()
PathText.Text = Dir1.Path + "\" + File1.filename
End Sub
Private Sub Form_Load()
With Tray
        .cbSize = Len(Tray)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        'The callback should be the mousemove event
        .MenuMsg = WM_MOUSEMOVE
        .TIcon = Me.Icon
        'This is the tooltip in the systray
        .msgTip = "WebServer 1.2" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, Tray

Oldheight = WebServer.height
Oldwidth = WebServer.width
B = 1
C = 0
Counter = 0
D = 0
Errorcnt = 0
Socks = "I"
List1.AddItem "Error 1"
List1.AddItem "Error 2"
List1.AddItem "Error 3"
List1.AddItem "Error 4"
List1.AddItem "Error 5"
List1.AddItem "Error 6"
Stufftwo = ""
Data = ""
Open "c:\webserver.cfg" For Binary Access Read As #1
        Setup = Space(LOF(1))
        Get #1, , Setup 'get the path for the index file out of the config
Close #1
        If Setup <> "" Then
            Path = Setup
            'get the index file out of the path
            HtmlIndexFile = Right(Path, Len(Path) - InStrRev(1, Path, "\") + 1)
            Path = Left(Setup, Len(Setup) - Len(HtmlIndexFile))
        Else
            'if there is no path then show frame1 so that the user can specify path
            Frame1.Visible = True
        End If
Randomize Timer
ConnectCnt = 0
ConnectionCounter = 0
    Disconnect.Enabled = False
    Disconnect1.Enabled = False
On Error GoTo pe
Open Path + HtmlIndexFile For Binary Access Read As #1
        Data = Space(LOF(1)) 'open the index file, set data to textbox
        Get #1, , Data
        HTML.Text = Data
        If Data = "" Then GoTo pe
        HtmlIndex = Data
        Close #1
        
        HTTP(0).LocalPort = 80 'set the winsock's local port
        A = 0
PathText.Text = Path + HtmlIndexFile
DftError 'run the sub to set the "404 errors"
GoTo 4
pe:
Frame1.Visible = True
4
End Sub

Private Sub HTTP_Close(Index As Integer)
'If the connecion is closed then unload or close the winsock
If Socks = "I" Then
  Unload HTTP(Index)
  Unload Timer1(Index)
Else
HTTP(Index).Close
End If
ConnectionCounter = ConnectionCounter + 1
cnt2 = ConnectionCounter
End Sub


Private Sub HTTP_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If Socks = "I" Then
    ConnectCnt = ConnectCnt + 1
    If ConnectCnt = 32000 Then ConnectCnt = 1
    Load HTTP(ConnectCnt) 'load another winsock
    Load Timer1(ConnectCnt) 'load a timer
    HTTP(ConnectCnt).Accept requestID 'accept the connection request
     
         'Checks if the clients IP is in the string, if so, it will not add it again, but if its not. It will add it.
        If InStr(1, IPList, HTTP(ConnectCnt).RemoteHostIP & " ", vbTextCompare) Then GoTo 24
            IPList = IPList & HTTP(ConnectCnt).RemoteHostIP & " "
24
If InStr(1, Blocked_IPs, HTTP(ConnectCnt).RemoteHostIP & " ", vbTextCompare) Then HTTP(ConnectCnt).Close: Exit Sub
Else
    For i% = 1 To Socks
        If HTTP(i%).State = 0 Then
            HTTP(i%).Accept requestID 'accept the connection request
    
         'Checks if the clients IP is in the string, if so, it will not add it again, but if its not. It will add it.
        If InStr(1, IPList, HTTP(ConnectCnt).RemoteHostIP & " ", vbTextCompare) Then GoTo 23
            IPList = IPList & HTTP(i%).RemoteHostIP & " "
23
            If InStr(1, Blocked_IPs, HTTP(i%).RemoteHostIP & " ", vbTextCompare) Then HTTP(i%).Close: Exit Sub
            GoTo 22
        End If
    Next i%
22
End If
    Counter = Counter + 1
    ConnectionCounter = ConnectionCounter + 1
    cnt1 = Counter
    cnt2 = ConnectionCounter
End Sub
Private Sub HTTP_DataArrival(Index As Integer, ByVal bytesTotal As Long)
        'Sends message if Ip matches the IP selected
        If HTTP(Index).RemoteHostIP = Label11.Caption And SendMsg = 1 Then SendMsg = 0: Data = MessageTxt.Text: Label15.Caption = "Sent": HTML.Text = MessageTxt.Text: GoTo 9
        
        HTTP(Index).GetData NewData, vbString 'Gets the data the client sent
        IPtxt.Text = NewData 'sets the data to textbox
        IPScmd.Caption = "IPs"
        A = 0
        If NewData = "" Then HTTP_SendComplete (Index): GoTo 3 'If no data was recieved: close winsock and exit the sub
        Request = Mid(NewData, 5, InStr(5, NewData, " ") - 5) 'Gets requested filename from clients data
        IPs = IPs + HTTP(ConnectCnt).RemoteHostIP & " => " & Request + Chr(13) + Chr(10) 'Add Clients IP to the list and Its request
            If Request <> "/" Then
                '/.. protection: if "/.." is found in request then server will disregard the request
                If InStr(1, Request, "/..", vbTextCompare) And SlashDot = 1 Then Data = "": GoTo 20
            On Error GoTo pe
            Open Path + Request For Binary Access Read As #1 'Opens Requested file
                On Error GoTo 0
                Data = Space(LOF(1))
                Get #1, , Data
                HTML.Text = Data
            Close #1
        If Data = "" Then Kill Path + Request: Call RandomError 'if nothing is in file, kill it and get "404 error"
20
    On Error Resume Next
    Stufftwo = Mid(Request, InStr(5, Request, "."))
1
If InStr(Data, LCase("<count>")) <> 0 Then 'if "<count>" is in the data then run the page counter rutine
       GoTo 6
Else
    GoTo 7
End If
6
Leftty = Mid(Data, 1, InStr(5, Data, "<count>") - 1) 'takes all data from the left of <count>
Righty = Mid(Data, InStr(5, Data, "<count>") + 7) 'takes all data from the right of <count>
        Open Path + Stufftwo + "Count" For Binary Access Read As #1 'Gets the # of hits from the file
            HitCount = Space(LOF(1))
            Get #1, , HitCount
        Close #1
If HitCount = "" Then HitCount = 1: GoTo 10
HitCount = HitCount + 1 'adds one to the hit counter
10
        Open Path + Stufftwo + "Count" For Binary Access Write As #1
            Put #1, , HitCount 'Saves HIts
            Close #1
Data = Leftty & "This page has been viewed " & HitCount & " times." & Righty 'Puts the data back togather
7
Stufftwo = LCase(Stufftwo)
Select Case Stufftwo 'sets the "Content-Type" so the client knows what kind of data is sent
    Case ".php"
        PagesViewed = PagesViewed + 1
        CodeType = "text/html"
    Case ".html"
        PagesViewed = PagesViewed + 1
        CodeType = "text/html"
    Case ".jpg"
        CodeType = "image/jpeg"
    Case ".gif"
        CodeType = "image/gif"
    Case ".zip"
        CodeType = "aplication/zip"
    Case ".exe"
        CodeType = "aplication/exe"
    Case ".mpg"
        CodeType = "movie/mpeg"
    Case Else
        CodeType = "text/html"
End Select
            'Sends data
            HTTP(Index).SendData "HTTP/1.0 200 OK" & vbCrLf & _
                                     "Content-Length: " & Len(Data) & vbCrLf & _
                                     "Content-Type: " & CodeType & vbCrLf & _
                                     vbCrLf & _
                                     Data
        BytesSentLb = BytesSentLb + Len(Data) 'Adds the amount of bytes to the total bytes sent
        Else 'if request is = to "/" the it sents the index file
        
        'If Data = "" Then Call RandomError
        Open Path + HtmlIndexFile For Binary Access Read As #1 'Opens index file
            Data = Space(LOF(1))
            Get #1, , Data
            HTML.Text = Data
        Close #1
                       
If InStr(Data, LCase("<count>")) <> 0 Then 'Checks for "<count>"
    GoTo 8
Else
    GoTo 9
End If
8
    Leftty = Mid(Data, 1, InStr(5, Data, "<count>") - 1) 'takes all data from the left of <count>
    Righty = Mid(Data, InStr(5, Data, "<count>") + 7) 'takes all data from the right of <count>
        Open Path + Stuff + "Count" For Binary Access Read As #1 'Gets the # of hits from the file
            HitCount = Space(LOF(1))
            Get #1, , HitCount
        Close #1
If HitCount = "" Then HitCount = 1: GoTo 11
HitCount = HitCount + 1 'adds one to the hit counter
11
        Open Path + Stuff + "Count" For Binary Access Write As #1
            Put #1, , HitCount 'Saves # of hits
        Close #1
Data = Leftty & "This page has been viewed " & HitCount & " times." & Righty
9
Stuff = Mid(HtmlIndexFile, InStr(5, HtmlIndexFile, "."))
Stuff = LCase(Stuff)
Select Case Stuff 'sets the "Content-Type" so the client knows what kind of data is sent
    Case ".php"
        PagesViewed = PagesViewed + 1
        CodeType = "text/html"
    Case ".html"
        PagesViewed = PagesViewed + 1
        CodeType = "text/html"
    Case ".jpg"
        CodeType = "image/jpeg"
    Case ".gif"
        CodeType = "image/gif"
    Case ".zip"
        CodeType = "aplication/zip"
    Case ".exe"
        CodeType = "aplication/exe"
    Case ".txt"
        PagesViewed = PagesViewed + 1
        CodeType = "text/plain"
    Case Else
        CodeType = "text/html"
End Select
                'Sends data to client
                HTTP(Index).SendData "HTTP/1.0 200 OK" & vbCrLf & _
                                     "Content-Length: " & Len(Data) & vbCrLf & _
                                     "Content-Type: " & CodeType & vbCrLf & _
                                     vbCrLf & _
                                     Data
        BytesSentLb = BytesSentLb + Len(Data) 'Adds the amount of bytes to the total bytes sent
        End If
        Exit Sub
GoTo 3
pe:
RandomError
GoTo 1
3
End Sub

Private Sub HTTP_SendComplete(Index As Integer)
'Once all data has been sent then the winsock needs to be closed
On Error Resume Next
HTTP(Index).Close
If Socks = "I" Then
    Unload HTTP(Index)
    Unload Timer1(Index)
End If
ConnectionCounter = ConnectionCounter - 1
cnt2 = ConnectionCounter
End Sub
Private Sub RandomError()
Errorcnt = Errorcnt + 1
Label8.Caption = Errorcnt
Error = Int(Rnd * 6) + 1 'selests a random "404 error"
Select Case Error
Case 1
Data = Error1
Case 2
Data = Error2
Case 3
Data = Error3
Case 4
Data = Error4
Case 5
Data = Error5
Case 6
Data = Error6
End Select
HTML.Text = Data
End Sub

Private Sub IPScmd_Click()
If A = 0 Then
    IPScmd.Caption = "Rqst"
    IPtxt.Text = IPs
    A = A + 1
Else
    IPScmd.Caption = "IPs"
    IPtxt.Text = NewData
    A = 0
End If
End Sub
Private Sub Okcmd_Click()
Connect1.Enabled = True
connect.Enabled = True
Path = PathText.Text
HtmlIndexFile = Right(Path, Len(Path) - InStrRev(1, Path, "\") + 1) 'Cuts the file and extention out of Path
Path = Left(Path, Len(Path) - Len(HtmlIndexFile))
Frame1.Visible = False
         Kill "c:\webserver.cfg"
         Open "c:\webserver.cfg" For Binary Access Write As #1
            Put #1, , PathText.Text 'Writes the config file
         Close #1
    On Error GoTo pe
        Open Path + HtmlIndexFile For Binary Access Read As #1
            Data = Space(LOF(1))
            Get #1, , Data 'Opens index file, and puts the Data into a textbox
            If Data = "" Then GoTo pe
            HTML.Text = Data
        Close #1
GoTo 2
pe: 'Error message if no data in file
MsgBox "Invalid path\file name, or no Data in file.", vbOKOnly, "Invalid path\file name"
Close #1
2
End Sub

Private Sub PathMenu_Click()
Frame1.Visible = True
Connect1.Enabled = False
connect.Enabled = False
Frame1.ZOrder (0)
End Sub

Function InStrRev(start As Long, string1 As String, string2 As String)
  'Cuts index filename from path
  Dim E As Long
  start = start + Len(string2)
  string1 = Left(string1, Len(string1) - start + Len(string2) + 1)
  E = 1
  While E <> 0
    start = E
    E = InStr(E + 1, string1, string2)
  Wend
  InStrRev = start
End Function

Private Sub SetSocks_Click()
Dim OldNum As Integer
OldNum = 0
If Socks <> "I" Then OldNum = Socks
If MaxSock.Text = "Infinite" Then
    If OldNum <> 0 Then
        For i% = 1 To OldNum 'Unloads all winsocks
            Unload HTTP(i%)
            Unload Timer1(i%)
        Next i%
    End If
    Socks = "I"
Else
If OldNum = MaxSock.Text Then GoTo 21
On Error Resume Next
    Socks = MaxSock.Text
    For i% = OldNum + 1 To MaxSock.Text 'Loads and additional winsocks
        Load HTTP(i%)
        Load Timer1(i%)
    Next i%
21
End If
Frame4.ZOrder 0
End Sub

Private Sub Showmnu_Click()
Me.Show
Showmnu.Enabled = False
Hidemnu.Enabled = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Frame4.ZOrder 0
End Sub


Private Sub Startmnu_Click()
HTTP(0).Listen
    Connected = "YES"
    Connect1.Enabled = False
    connect.Enabled = False
    Reset.Enabled = True
    Disconnect.Enabled = True
    Disconnect1.Enabled = True
Command10.Enabled = False
Command11.Enabled = False
Command12.Enabled = False
Command13.Enabled = False
Startmnu.Enabled = False
Stopmnu.Enabled = True
SetSocks.Enabled = False
Command15.Enabled = False
PathMenu.Enabled = False
End Sub

Private Sub Stats_Click()
Text2.Text = ""
AllIPs = IPList
On Error GoTo pe
Do 'Adds IPs to list, and makes sure there Isnt any Dubbles
OneIP = Mid(AllIPs, 1, InStr(1, AllIPs, " ") - 1)
AllIPs = Mid(AllIPs, InStr(1, AllIPs, " ") + 1)
'If InStr(1, Text2.Text, OneIP, vbTextCompare) Then GoTo 17
Text2.Text = Text2.Text & OneIP & Chr(13) & Chr(10)
'17
Loop Until AllIPs = ""
GoTo 16
pe:
Text2.Text = Text2.Text & AllIPs
16
Frame3.Visible = True
Frame3.ZOrder (0) 'Sets frame3 on top
Frame3.Refresh
Picture2.CurrentX = 0
Picture2.CurrentY = 0
Picture2.Print "Stats"
End Sub

Private Sub Stopmnu_Click()
HTTP(0).Close
    Open Path + HtmlIndexFile For Binary Access Read As #1 'Opens index file
        Data = Space(LOF(1))
        Get #1, , Data
        HTML.Text = Data 'Sets file data to textbox
        Close #1
    NewData = ""
    IPtxt = ""
    HTML.Text = Data
    connect.Enabled = True
    Connect1.Enabled = True
    Reset.Enabled = False
    Disconnect.Enabled = False
    Disconnect1.Enabled = False
    Stopmnu.Enabled = False
    Startmnu.Enabled = True
    Command10.Enabled = True
    Command11.Enabled = True
    Command12.Enabled = True
    Command13.Enabled = True
    SetSocks.Enabled = True
    Command15.Enabled = True
On Error Resume Next
If Socks = "I" Then
    For i% = 1 To ConnectCnt 'Unload winsocks and timers
      HTTP(i%).Close
      Unload HTTP(i%)
      Unload Timer1(i%)
    Next i%
End If
    ConnectCnt = 0
    Counter = 0
    ConnectionCounter = 0
    cnt1 = Counter
    cnt2 = ConnectionCounter
PathMenu.Enabled = True
End Sub

Private Sub Tabs_Click()
List2.Clear
List4.Clear
AllIPs = IPList
On Error GoTo 18

Do 'Adds IPs to list, and makes sure there Isnt any Dubbles
OneIP = Mid(AllIPs, 1, InStr(1, AllIPs, " ") - 1)
AllIPs = Mid(AllIPs, InStr(1, AllIPs, " ") + 1)
List2.AddItem OneIP
List4.AddItem OneIP
    'For i% = 1 To (List2.ListCount + 1)
     '   If List3.List(List3.ListIndex = i%) <> OneIP And List2.List(List2.ListIndex = i%) <> OneIP Then List2.AddItem OneIP: List4.AddItem OneIP
     '   If List4.List(List4.ListIndex = i%) <> OneIP Then List4.AddItem OneIP
    'Next i%
19
Loop Until AllIPs = ""
GoTo 18
18
Frame4.Visible = True
Frame4.ZOrder (0) 'Sets frame4 on top
Frame4.Refresh

Picture1.CurrentX = 0
Picture1.CurrentY = 0
Picture1.Print "Security"
End Sub

Private Sub Text1_Click()
Frame4.ZOrder 0
End Sub

Private Sub Timeoutcmd_Click()
If B = 0 Then
B = 1
C = 1
Dissable.Caption = "Enable"
Timeoutcmd.Caption = "Off"
Label3.Caption = "Connection Timeout Disabled."
Timer1(0).Enabled = False
Else
Dissable.Caption = "Disable"
Label3.Caption = "Connection Timeout Enabled."
Timeoutcmd.Caption = "On"
Timer1(0).Enabled = True
C = 0
B = 0
End If
End Sub

Private Sub Timer1_Timer(Index As Integer) 'Connection timeout
If Index <> 0 And C = 0 And Socks = "I" Then
    Unload HTTP(Index)
    Unload Timer1(Index)
    ConnectionCounter = ConnectionCounter - 1
Else
    If Index <> 0 Then HTTP(Index).Close
End If
End Sub
Private Sub DftError()
Error1 = "<html><head><title>HTTP Error 404</title></head><body><h1>Page cannot be displayed.</h1><p>Try <a href='javascript:location.reload()'>refreshing</a> the page.</p><p>If That doesn't work go <a href='javascript:history.back(1)'>back</a> and try another link.</p></body></html>"
Error1A = Error1
Error2 = "<html><head><title>HTTP Error 404</title></head><body><h1>Umm, i can't seem to find that there file.</h1><p>Try <a href='javascript:location.reload()'>refreshing</a> the page.</p><p>If That doesn't work go <a href='javascript:history.back(1)'>back</a> and try another link.</body></html>"
Error2A = Error2
Error3 = "<html><head><title>HTTP Error 404</title></head><body><h1>No such file....</h1><h1>What are you trying to do anyways?</h1></body></html>"
Error3A = Error3
Error4 = "<html><head><title>HTTP Error 404</title></head><body><h1>Sorry dude...</h1>I can't seem to find a file with that name. Why don't you just try it <a href = 'javascript:location.reload()'>again</a>.</body></html>"
Error4A = Error4
Error5 = "<html><head><title>HTTP Error 404</title></head><body><h1>Now you've done it!!</h1><h1>Okay, it might not have been your fault, but I can't seem to find that file.</h1><p>Try <a href='javascript:location.reload()'>refreshing</a> the page.</p><p>If That doesn't work go <a href='javascript:history.back(1)'>back</a> and try another link.</p></body></html>"
Error5A = Error5
Error6 = "<html><head><title>HTTP Error 404</title></head><body><h1>You broke it!!</h1><h3>well acctualy I just cant't find the file you have requested. Don't worry 'bout it and try <a href = 'javascript:location.reload()'>again</a>.</h3></body></html>"
Error6A = Error6
End Sub

