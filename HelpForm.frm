VERSION 5.00
Begin VB.Form HelpForm 
   Caption         =   "Help"
   ClientHeight    =   4815
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5070
   Icon            =   "HelpForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   4815
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "HelpForm.frx":030A
      Top             =   0
      Width           =   5055
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Site 
         Caption         =   "My Site"
      End
      Begin VB.Menu Close 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "HelpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HelpText As String
Private Sub Close_Click()
Unload Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
Text1.Width = HelpForm.ScaleWidth
Text1.Height = HelpForm.ScaleHeight
End Sub

Private Sub Site_Click()
Shell "start www.learnelectronics.f2s.com"
    'Opens internet explorer and sets url to www.learnelectronics.f2s.com
End Sub
