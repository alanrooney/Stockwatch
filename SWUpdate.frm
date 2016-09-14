VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmSWUpdate 
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   540
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   780
      Width           =   1410
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

Private Sub Form_Load()
Dim varBkgs As Variant
Dim sV As String
Dim sVer As String

    Inet1.URL = "http://www.glick.ie"
    Inet1.UserName = "glickie"
    Inet1.Password = "qapl10wsok"
    Inet1.RequestTimeout = 120
    Inet1.Protocol = icFTP
    
    
    Inet1.Execute Inet1.URL, "CD SWUpdates"
    Stop
    
    
    Inet1.Execute Inet1.URL, "DIR"
    Stop
        
        
        
        varBkgs = Inet1.GetChunk(1024, icString)
        sV = varBkgs

    sVer = Mid(sV, InStr(sV, "Ver") + 3, 3)
    Text1 = sVer

End Sub
