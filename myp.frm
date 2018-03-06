VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form myp 
   Caption         =   "My own tutorials"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   8670
   LinkTopic       =   "Form2"
   ScaleHeight     =   7035
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton st 
      Caption         =   "&Save tutorial"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   6480
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   10186
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "myp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String

Private Sub Form_Load()
Dim tutfb As String

'On Error GoTo z

  If FileExists(App.Path & "\Data\template.htutt") Then
  FileCopy (App.Path & "\Data\template.htutt"), App.Path & "\My tutorials\unnamed.htut"
  Else
  MsgBox "Files are missing, please reinstall applicaion", vbCritical, "Files are missing"
  GoTo zz
  End If
 
  If FileExists(App.Path & "\My tutorials\unnamed.htut") Then
  tutfb = App.Path & "\My tutorials\unnamed.htut"
  gstrConnectionString2 = ("provider=microsoft.jet.oledb.4.0;data source= " & tutfb)
  Else
  MsgBox "Unknown error, make sure disk is not full or write protected", vbCritical
  GoTo zz
  End If
  
  Set adoconn = Nothing
adoconn.Open gstrConnectionString2


  Set rs = Nothing
    str = "puzzles"


  rs.Open str, adoconn, adOpenForwardOnly, adLockReadOnly
  
  
 Set rs = New ADODB.Recordset
   rs.CursorLocation = adUseClient
   
   ' Add columns to the Recordset
   rs.Fields.Append "Key", adInteger
   rs.Fields.Append "Field1", adVarChar, 40, adFldIsNullable
   rs.Fields.Append "Field2", adDate

   ' Open the Recordset
rs.Open str, adoconn, adOpenStatic, adLockBatchOptimistic

   
   ' Populate the Data in the DataGrid
   Set grid.DataSource = rs


GoTo zz
z:

MsgBox "Error, Database is corrupted, can't continue", vbCritical, Error
zz:

End Sub

