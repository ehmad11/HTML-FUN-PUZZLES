VERSION 5.00
Begin VB.Form Learn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About <some tag>"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6630
   Icon            =   "Learn.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6630
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Default         =   -1  'True
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Learn more on w3c Schools"
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   4320
      Width           =   2895
   End
   Begin VB.TextBox itext 
      Height          =   2535
      Left            =   1320
      TabIndex        =   1
      Text            =   "Info about this tag"
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "Information:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label labeltag 
      AutoSize        =   -1  'True
      Caption         =   "<some tag>"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1965
   End
End
Attribute VB_Name = "Learn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim adoconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String




Private Sub Command1_Click()
Dim ret As Long, theWebSite As String
'theWebSite = "http://w3c.org/"

theWebSite = "http://www.google.com/search?sitesearch=www.w3schools.com&as_q=" & labeltag.Caption & "&x=0&y=0"

ret = ShellExecute(Me.hwnd, "open", theWebSite, vbNullString, vbNullString, 3)
If ret < 32 Then MsgBox "There was an error when trying to open a default browser", vbCritical, "Error"

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()

Me.Caption = "Learn about " & learnq
labeltag.Caption = CStr(learnq)

'On Error GoTo z

'gstrConnectionString = ("provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\tags_help.mdb")


Set adoconn = Nothing
adoconn.Open gstrConnectionString

Set rs = Nothing
'str = "select * from puzzles where tag=" & Key
str = CStr("select * from puzzles where tag LIKE '" & learnq & "%'") ' "select * from cafe where Date=" & CStr(Date)

rs.Open str, adoconn, adOpenForwardOnly, adLockOptimistic
    

    'Dim intX    As Integer
    'For intX = 1 To 100
    'List1.AddItem "Item" & intX
    'Next


rs.MoveFirst

If rs(2) = "" Then

Else
itext.Text = rs(2)
End If

While Not rs.EOF


rs.MoveNext
Wend





GoTo zz

z:
   MsgBox "Info about this tag not found"
    Unload Me
    
zz:




End Sub
