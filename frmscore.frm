VERSION 5.00
Begin VB.Form frmscore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Scores"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7185
   Icon            =   "frmscore.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   7185
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2520
      TabIndex        =   1
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "High Scores"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "frmscore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String


Private Sub Command1_Click()
Me.Visible = False

End Sub

Function letsdo()

Label2.Caption = "Name"
Label3.Caption = "Score"

On Error GoTo z

gstrConnectionString3 = ("provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\users.dat")

Set adoconn = Nothing
adoconn.Open gstrConnectionString3

Set rs = Nothing
'str = "scores"
str = "select * from scores order by score desc"

rs.Open str, adoconn, adOpenForwardOnly, adLockOptimistic
    
If Not rs.EOF Then rs.MoveFirst

While Not rs.EOF

Label2.Caption = Label2.Caption & vbCrLf & rs(1)
Label3.Caption = Label3.Caption & vbCrLf & rs(2)

rs.MoveNext
Wend

    
z:

End Function

Private Sub Form_Load()
letsdo

letsdo

End Sub
