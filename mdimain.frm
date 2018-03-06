VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm mdimain 
   BackColor       =   &H00000000&
   Caption         =   "HTML FUN PUZZLES"
   ClientHeight    =   7650
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   9450
   Icon            =   "mdimain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdimain.frx":6852
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cmd 
      Left            =   0
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   """htut"""
      DialogTitle     =   "Open tutorial file"
      Filter          =   "Tutorial Files|*.htut"
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnunew 
         Caption         =   "&New Game"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuselect 
         Caption         =   "Select tutorial"
      End
      Begin VB.Menu mnudash 
         Caption         =   "-"
      End
      Begin VB.Menu mnudownload 
         Caption         =   "&Download More Tutorials"
      End
      Begin VB.Menu mnudashb 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnumyt 
         Caption         =   "Run HTML EDITOR"
         Visible         =   0   'False
      End
      Begin VB.Menu mnudashc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "mdimain"
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
   

Private Sub MDIForm_Initialize()

On Error GoTo z

If FileExists(App.Path & "\back.jpg") = True Then
Me.Picture = LoadPicture((App.Path & "\back.jpg"))
Me.Show
End If

z:

End Sub

Private Sub MDIForm_Load()


Dim toolfrm As New frmmain
toolfrm.Show



End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Unload frmAbout
End Sub

Private Sub mnuabout_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub mnudownload_Click()
ShellExecute Me.hwnd, _
        vbNullString, _
        "https://go.ehmad11.com/hfp", _
        vbNullString, _
        "c:\", _
        SW_SHOWNORMAL

End Sub

Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub mnumyt_Click()

Shell (App.Path & "\Editor.exe")
'Dim frmt As New myp
'frmt.Show

End Sub

Private Sub mnuselect_Click()
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
