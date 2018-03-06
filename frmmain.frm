VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmmain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "HTML FUN PUZZLES"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   ControlBox      =   0   'False
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "frmmain.frx":6852
   ScaleHeight     =   4515
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Music"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   1200
      TabIndex        =   8
      Top             =   4680
      Width           =   8295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Puzzles"
      ForeColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   480
      TabIndex        =   7
      Top             =   2160
      Width           =   8415
      Begin VB.CommandButton Command4 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   6240
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Download more tutorials"
         Height          =   375
         Left            =   4200
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&View high scores"
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Select a tutorial"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Make your own tutorial"
      Height          =   1095
      Left            =   2760
      TabIndex        =   4
      Top             =   5640
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   """htut"""
      DialogTitle     =   "Open tutorial file"
      Filter          =   "Tutorial Files|*.htut"
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   735
      Left            =   480
      TabIndex        =   9
      Top             =   3480
      Width           =   8415
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   14843
      _cy             =   1296
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Learn HTML by solving puzzles."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1320
      Width           =   6135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "HTML FUN PUZZLES"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   600
      TabIndex        =   6
      Top             =   600
      Width           =   7215
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib _
   "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hwnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long
   Private SW_SHOWNORMAL
   
Private Sub Command1_Click()

On Error Resume Next

cmd.InitDir = App.Path & "\"

cmd.ShowOpen


If cmd.filename = "" Then
Else

tutf = cmd.filename
Dim Form As New Form1
Form.Show
'Unload Me
End If




End Sub

Private Sub Command3_Click()


ShellExecute Me.hwnd, _
        vbNullString, _
        "http://go.ehmad11.com/hfp", _
        vbNullString, _
        "c:\", _
        SW_SHOWNORMAL


End Sub

Private Sub Command4_Click()
Unload mdimain

End Sub

Private Sub Command5_Click()

Unload frmscore

Load frmscore
frmscore.Show

End Sub

Private Sub Form_Load()

Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

wmp.URL = App.Path & "\playlist.m3u"

End Sub

