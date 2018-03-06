VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTML FUN PUZZLE"
   ClientHeight    =   8910
   ClientLeft      =   4740
   ClientTop       =   2655
   ClientWidth     =   15105
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   15105
   Begin VB.CommandButton Command6 
      Caption         =   "&Check progress"
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox txtbb 
      Height          =   2655
      Left            =   8640
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "Form1.frx":6852
      Top             =   2040
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Timer Timer 
      Interval        =   1
      Left            =   360
      Top             =   2520
   End
   Begin VB.ListBox List1 
      DragIcon        =   "Form1.frx":6858
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5310
      Left            =   240
      TabIndex        =   13
      Top             =   600
      Width           =   6675
   End
   Begin VB.Frame Frame4 
      Caption         =   "Options"
      Height          =   1935
      Left            =   7080
      TabIndex        =   11
      Top             =   6720
      Width           =   2535
      Begin VB.CommandButton Command3 
         Caption         =   "&Select another puzzle"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Restart"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Shuffle Again"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Time Left"
      Height          =   1935
      Left            =   9600
      TabIndex        =   9
      Top             =   6720
      Width           =   2535
      Begin VB.Label Label1 
         Caption         =   "Click here to stop/unstop time"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label lt 
         Alignment       =   2  'Center
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "About this tutorial"
      Height          =   2175
      Left            =   240
      TabIndex        =   7
      Top             =   6600
      Width           =   6615
      Begin VB.CommandButton Command5 
         Caption         =   "&Read out"
         Height          =   255
         Left            =   5520
         TabIndex        =   5
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Text            =   "Form1.frx":769A
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   11520
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "Form1.frx":76A6
      Top             =   720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox List2 
      DragIcon        =   "Form1.frx":76AC
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   -480
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Submit / &Validate Solution"
      Default         =   -1  'True
      Height          =   1455
      Left            =   12360
      TabIndex        =   4
      Top             =   6960
      Width           =   2415
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   5775
      Left            =   7080
      TabIndex        =   12
      Top             =   600
      Width           =   7695
      ExtentX         =   13573
      ExtentY         =   10186
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label3 
      Caption         =   "Preview of final page ( what it should look like)."
      Height          =   255
      Left            =   7080
      TabIndex        =   16
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Arrange tags to make similiar web page on right of screen."
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   240
      Width           =   4335
   End
   Begin VB.Menu mnuinfo 
      Caption         =   "Info"
      Visible         =   0   'False
      Begin VB.Menu mnulearn 
         Caption         =   "Learn about this tag"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cont, won As Boolean

Dim timel, jugaar As Integer

Dim adoconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String


Private Declare Function SendMessage _
                         Lib "user32" _
                         Alias "SendMessageA" _
                         (ByVal hwnd As Long, _
                         ByVal wMsg As Long, _
                         ByVal wParam As Long, _
                         lParam As Any) _
As Long

Private mintDragIndex       As Integer

Private Sub Command1_Click()
Dim i, rand As Integer
Dim str1 As String

For i = 0 To List1.ListCount - 1
 rand = Int(Rnd * List1.ListCount)
  
  str1 = List1.List(i)
  List1.List(i) = List1.List(rand)
  List1.List(rand) = str1
Next i


End Sub

Private Sub Command2_Click()

Dim solved As Boolean
Dim x As Integer
x = 0

solved = True

For x = 0 To List1.ListCount - 1

If List1.List(x) = List2.List(x) Then
Else
solved = False
End If

'x = x + 1
Next

If solved = True Then

cont = False

won = True


MsgBox "Hooooray, u have done it :D"

If won = True Then submitscore (lt.Caption)
If (cont = False And won = False) Then MsgBox "But timer was disabled :P"


Unload Me

Else
MsgBox "Still not solved"
End If

setpre

End Sub

Private Sub Command3_Click()

 'If MsgBox("Are you sure you want to exit current puzzle", vbYesNo, "Confirm exit") = vbYes Then
 

'Load frmmain
'frmmain.Show

Unload Me

'End If

End Sub

Private Sub Command4_Click()
gettinfo
 Command1_Click

End Sub

Private Sub Command5_Click()

On Error Resume Next

Dim frmi As New frminfo
frmi.Show
frmi.Text1.Text = CStr(Text2.Text)

End Sub

Private Sub Command6_Click()
setpre
End Sub

Private Sub Command7_Click()

End Sub

'-----------------------------------------------------------------------------
Private Sub Form_Load()
'-----------------------------------------------------------------------------

jugaar = 0

cont = True
won = False

Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
    
    
On Error GoTo z

'gstrConnectionString = ("provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\db.mdb")
gstrConnectionString = ("provider=microsoft.jet.oledb.4.0;data source= " & tutf)     '" & App.Path & "\db.mdb")

Set adoconn = Nothing
adoconn.Open gstrConnectionString

Set rs = Nothing
str = "puzzles"

rs.Open str, adoconn, adOpenForwardOnly, adLockOptimistic
    

    'Dim intX    As Integer
    'For intX = 1 To 100
    'List1.AddItem "Item" & intX
    'Next


rs.MoveFirst
Text1.Text = ""


While Not rs.EOF

List1.AddItem rs(1)
List2.AddItem rs(1)

Text1.Text = Text1.Text & rs(1) & vbCrLf  'List2.List(x) & vbCrLf

rs.MoveNext
Wend

gettinfo

Command1_Click

If FileExists(App.Path & "\solution.html") Then
Kill App.Path & "\solution.html"
End If

Open App.Path & "\solution.html" For Append As #1
Print #1, Text1.Text    ' List2.List(x) '& vbCrLf   'Text1
Close #1

wb.Navigate App.Path & "\solution.html"

GoTo zz

z:
   MsgBox "Invalid tutorial", vbCritical, Error
   Load frmmain
   frmmain.Show
    Unload Me
    
zz:


End Sub

Private Sub Form_Unload(Cancel As Integer)
If (won <> True) Then


If (MsgBox("Are you sure you want to exit", vbYesNo, "Confirm exit") = vbYes) Then

Else
Cancel = 2

End If

End If

End Sub

Private Sub Label1_Click()
cont = Not cont

End Sub

'-----------------------------------------------------------------------------
Private Sub List1_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  Y As Single)
'-----------------------------------------------------------------------------
If cont = True Then
mintDragIndex = ListRowCalc(List1, Y)
    List1.Drag
    End If
    
End Sub

'-----------------------------------------------------------------------------
Private Sub List1_DragOver(Source As Control, _
                           x As Single, _
                           Y As Single, _
                           State As Integer)
'-----------------------------------------------------------------------------
If cont = True Then List1.ListIndex = ListRowCalc(List1, Y)
End Sub

'-----------------------------------------------------------------------------
Private Sub List1_DragDrop(Source As Control, _
                           x As Single, _
                           Y As Single)
'-----------------------------------------------------------------------------
    
    
If cont = True Then ListRowMove Source, mintDragIndex, ListRowCalc(Source, Y)
End Sub

'-----------------------------------------------------------------------------
Public Function ListRowCalc(pobjLB As ListBox, ByVal Y As Single) As Integer
'-----------------------------------------------------------------------------
           
    Const LB_GETITEMHEIGHT = &H1A1
    
    Dim intItemHeight   As Integer
    Dim intRow          As Integer
    
    intItemHeight = SendMessage(pobjLB.hwnd, LB_GETITEMHEIGHT, 0, 0)
    
    intRow = ((Y / Screen.TwipsPerPixelY) \ intItemHeight) + pobjLB.TopIndex
    
    If intRow < pobjLB.ListCount - 1 Then
        ListRowCalc = intRow
    Else
        ListRowCalc = pobjLB.ListCount - 1
    End If
                 
End Function

'-----------------------------------------------------------------------------
Public Sub ListRowMove(pobjLB As ListBox, _
                       ByVal pintOldRow As Integer, _
                       ByVal pintNewRow As Integer)
'-----------------------------------------------------------------------------
                       
    Dim strSavedItem    As String
    Dim intX            As Integer

    If pintOldRow = pintNewRow Then Exit Sub
    
    strSavedItem = pobjLB.List(pintOldRow)
    
    If pintOldRow > pintNewRow Then
        For intX = pintOldRow To pintNewRow + 1 Step -1
            pobjLB.List(intX) = pobjLB.List(intX - 1)
        Next intX
    Else
        For intX = pintOldRow To pintNewRow - 1
            pobjLB.List(intX) = pobjLB.List(intX + 1)
        Next intX
    End If
    
    pobjLB.List(pintNewRow) = strSavedItem

End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

If jugaar > 1 Then


learnq = List1.Text

Dim lrn As New Learn
lrn.Show



'PopupMenu mnuinfo

If Button = 2 Then
     PopupMenu mnuinfo
 End If

Else
jugaar = 2
End If

End Sub

Private Sub mnulearn_Click()
learnq = List1.Text

Dim lrn As New Learn
lrn.Show

End Sub

Private Sub Option1_Click()
wb.Visible = False
List1.Visible = True

End Sub

Private Sub Option2_Click()
wb.Visible = True
List1.Visible = False

End Sub

Function gettinfo()

On Error GoTo z

'gstrConnectionString = ("provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\db.mdb")
gstrConnectionString = ("provider=microsoft.jet.oledb.4.0;data source= " & tutf)     '" & App.Path & "\db.mdb")

Set adoconn = Nothing
adoconn.Open gstrConnectionString

Set rs = Nothing
str = "info"

rs.Open str, adoconn, adOpenForwardOnly, adLockOptimistic


If Not rs.EOF Then
rs.MoveFirst
Text2.Text = rs(1)
lt.Caption = rs(2)
Else
Text2.Text = "No info found about this tutorial"
End If


GoTo zz

z:
   MsgBox "Invalid tutorial", vbCritical, Error
   Load frmmain
   frmmain.Show
    Unload Me
    
zz:

End Function

Private Sub Timer_Timer()
On Error Resume Next
Timer.Interval = 1000

timel = CInt(lt)

If timel = 0 Then
gettinfo
 Command1_Click
MsgBox "Sorry, you just lost the game"

'Unload Me
Else
List1.Enabled = True

If cont = True Then
timel = timel - 1

lt.Caption = CStr(timel)
Else
'List1.Enabled = False

End If

End If
End Sub


Function setpre()

Dim x As Integer

On Error Resume Next


x = 1
List1.Enabled = False

txtbb.Text = ""


For x = 0 To List1.ListCount

txtbb.Text = txtbb.Text & List1.List(x) & vbCrLf  'List2.List(x) & vbCrLf
'MsgBox x

'x = x + 1
Next

List1.Enabled = True


If FileExists(App.Path & "\solution_urs.html") Then
Kill App.Path & "\solution_urs.html"
End If

Open App.Path & "\solution_urs.html" For Append As #1
Print #1, txtbb.Text    ' List2.List(x) '& vbCrLf   'Text1
Close #1

Load frmsolurs
frmsolurs.wb.Navigate App.Path & "\solution_urs.html"
frmsolurs.Show

GoTo zz

z:
    
zz:


End Function
