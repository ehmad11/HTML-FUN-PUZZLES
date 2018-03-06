Attribute VB_Name = "SystemAPI"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Dim adoconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String

Global tutf, learnq As String

Global gstrConnectionString, gstrConnectionString2, gstrConnectionString3, passstr As String

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

' The Windows directory.

Function WindowsDirectory() As String
    Dim buffer As String * 512, length As Long
    length = GetWindowsDirectory(buffer, Len(buffer))
    WindowsDirectory = Left$(buffer, length)
End Function

' The System directory

Function SystemDirectory() As String
    Dim buffer As String * 512, length As Long
    length = GetSystemDirectory(buffer, Len(buffer))
    SystemDirectory = Left$(buffer, length)
End Function

' the Temp directory

Function TemporaryDirectory() As String
    Dim buffer As String * 512, length As Long
    length = GetTempPath(Len(buffer), buffer)
    TemporaryDirectory = Left$(buffer, length)
End Function

' The user's name

Function UserName() As String
    Dim buffer As String * 512, length As Long
    If GetUserName(buffer, Len(buffer)) Then
        length = InStr(buffer, vbNullChar) - 1
        UserName = Left$(buffer, length)
    End If
End Function

' The name of the computer

Property Get ComputerName() As String
    Dim buffer As String * 512, length As Long
    length = Len(buffer)
    If GetComputerName(buffer, length) Then
        ' returns non-zero if successful, and modifies the length argument
        ComputerName = Left$(buffer, length)
    End If
End Property

' Return True if running under Windows NT, False if running under Win9x.

Function WindowsNT() As Boolean
    ' Running under NT if the sign bit is off.
    WindowsNT = (GetVersion >= 0)
End Function

' Windows version as a string.

Function WindowsVersion() As String
    Dim os As OSVERSIONINFO, ver As String
    ' The function expects the UDT size in its first element.
    os.dwOSVersionInfoSize = Len(os)
    GetVersionEx os
    WindowsVersion = os.dwMajorVersion & "." & Right$("0" & Format$(os.dwMinorVersion), 2)
End Function

' Windows Build number.

Function WindowsBuildNumber() As Long
    Dim os As OSVERSIONINFO, ver As String
    ' The function expects the UDT size in its first element.
    os.dwOSVersionInfoSize = Len(os)
    GetVersionEx os
    WindowsBuildNumber = os.dwBuildNumber
End Function

Function FileExists(ByVal filename As String) As Boolean
    On Error Resume Next
    FileExists = (Dir$(filename) <> "")
End Function

Function submitscore(score As String)

Dim name As String

name = InputBox("Please enter your name")

On Error GoTo z

gstrConnectionString3 = ("provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\users.dat")

'Set adoconn = Nothing
adoconn.Open gstrConnectionString3

Set rs = Nothing
str = "scores"
rs.Open str, adoconn, adOpenForwardOnly, adLockOptimistic
    
rs.AddNew
rs("Name") = CStr(name)
rs("Score") = CStr(score)


rs.Update

    
z:

'Load frmscore
'frmscore.Show


End Function
