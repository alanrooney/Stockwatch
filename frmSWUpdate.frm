VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmSWUpdate 
   Caption         =   "Form1"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1395
      TabIndex        =   1
      Top             =   285
      Width           =   1680
   End
   Begin VB.TextBox Text1 
      Height          =   1545
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmSWUpdate.frx":0000
      Top             =   780
      Width           =   2925
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   255
      Top             =   2220
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      RemoteHost      =   "www.glick.ie"
      URL             =   "http://www.glick.ie/SWUpdates"
      Document        =   "/SWUpdates"
   End
End
Attribute VB_Name = "frmSWUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

        
Private Declare Function InternetOpen _
    Lib "wininet.dll" Alias "InternetOpenA" ( _
        ByVal sAgent As String, _
        ByVal lAccessType As Long, _
        ByVal sProxyName As String, _
        ByVal sProxyBypass As String, _
        ByVal lFlags As Long) As Long

Private Declare Function InternetConnect _
    Lib "wininet.dll" Alias "InternetConnectA" ( _
        ByVal hInternetSession As Long, _
        ByVal sServerName As String, _
        ByVal nServerPort As Integer, _
        ByVal sUsername As String, _
        ByVal sPassword As String, _
        ByVal lService As Long, _
        ByVal lFlags As Long, _
        ByVal lContext As Long) As Long

Private Declare Function InternetCloseHandle _
    Lib "wininet.dll" (ByVal hInet As Long) As Integer

Private Declare Function FtpCommand _
    Lib "wininet.dll" Alias "FtpCommandA" ( _
        ByVal hConnect As Long, _
        ByVal fExpectResponse As Boolean, _
        ByVal dwFlags As Long, _
        ByVal lpszCommand As String, _
        dwContext As Long, _
        phFtpCommand As Long) As Boolean

Private Declare Function InternetReadFile _
    Lib "wininet.dll" ( _
        ByVal hConnect As Long, _
        ByVal lpBuffer As String, _
        ByVal dwNumberOfBytesToRead As Long, _
        lpdwNumberOfBytesRead As Long) As Boolean

'=======================================================
' End Declarations - Begin Procedures
'=======================================================

Private Function TrimNull$(ByVal s$)
    Dim pos%
    
    pos = InStr(s, Chr$(0))
    If pos Then s = Left$(s, pos - 1)
    
    TrimNull = Trim$(s)
End Function

Private Sub Command1_Click()
Dim varBkgs As Variant
Dim sV As String
Dim sVer As String
    Const NUMBYTES& = 1020
    Dim hOpen&, hConn&, hOutConn&, buffer$, bytesRead&
    
    hOpen = InternetOpen( _
                scUserAgent, _
                INTERNET_OPEN_TYPE_DIRECT, _
                vbNullString, _
                vbNullString, 0)
    
    DoEvents
    If hOpen = 0 Then Exit Sub
    
    'Note: This is coded for the anonymous Microsoft site
    '  See the Declare above for replacing the three
    '  parameters with your site, username and password.
    hConn = InternetConnect( _
                hOpen, _
                "www.glick.ie", _
                INTERNET_INVALID_PORT_NUMBER, _
                "glickie", _
                "qapl10wsok", _
                INTERNET_SERVICE_FTP, _
                INTERNET_FLAG_PASSIVE, 0)
    
    DoEvents
    If hConn = 0 Then GoTo out2
   
' Command to change a Directory
    FtpCommand hConn, _
                False, _
                FTP_TRANSFER_TYPE_ASCII, _
                "CWD SWUpdates", _
                0, _
                hOutConn
    
' Command to LIST the Directory
    FtpCommand hConn, _
                True, _
                FTP_TRANSFER_TYPE_ASCII, _
                "LIST", _
                0, _
                hOutConn
    
    DoEvents
    If hOutConn = 0 Then GoTo out1
    
    Text1 = ""
    buffer = Space$(NUMBYTES + 4)

    Do
        InternetReadFile _
                hOutConn, _
                buffer, _
                NUMBYTES, _
                bytesRead
        
        If bytesRead = 0 Then Exit Do
        sV = sV & TrimNull(buffer)
    Loop

    InternetCloseHandle hOutConn


    Text1 = Mid(sV, InStr(sV, "Ver") + 3, 3)


'        ' Download file
        If FtpGetFile(hConn, "StockwatchTestVer401.txt", "StockwatchTestVer401.txt", False, FILE_ATTRIBUTE_ARCHIVE, _
            FTP_TRANSFER_TYPE_UNKNOWN, 0&) Then
            ' Success
            
            'ErrLog "clsFTP.DownloadFile: Download Succeeded. Time:" & CStr(Now), 4
        Else
            ' Raise an error
            ErrLog "clsFTP.DownloadFile: Download Failed. " & Err.Number & ":" & _
                                                            Err.Description, 1
        End If

out1: InternetCloseHandle hConn

out2: InternetCloseHandle hOpen


End Sub


Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    Move 1000, 1000, 5400, 3600
    Text1.Move 120, 120, 5000, 2500
    Command1.Move 2000, 2760, 1200, 375
End Sub


'''Dim varBkgs As Variant
'''Dim sV As String
'''Dim sVer As String
'''
'''    Inet1.URL = "http://www.glick.ie"
'''    Inet1.UserName = "glickie"
'''    Inet1.Password = "qapl10wsok"
'''    Inet1.RequestTimeout = 120
'''    Inet1.Protocol = icFTP
'''
'''
'''    Inet1.Execute Inet1.URL, "CD SWUpdates"
'''    Stop
'''
'''
'''    Inet1.Execute Inet1.URL, "DIR"
'''    Stop
'''
'''
'''
'''        varBkgs = Inet1.GetChunk(1024, icString)
'''        sV = varBkgs
'''
'''    sVer = Mid(sV, InStr(sV, "Ver") + 3, 3)
'''    Text1 = sVer
'''

