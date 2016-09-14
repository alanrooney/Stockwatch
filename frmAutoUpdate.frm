VERSION 5.00
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmAutoUpdate 
   BorderStyle     =   0  'None
   ClientHeight    =   2085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   Picture         =   "frmAutoUpdate.frx":0000
   ScaleHeight     =   2085
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MyCommandButton.MyButton btnClose 
      Height          =   255
      Left            =   4065
      TabIndex        =   0
      Top             =   135
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   450
      BackColor       =   13748165
      ForeColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   9
      BackColorDown   =   3968251
      BackColorOver   =   15200231
      BackColorFocus  =   13748165
      BackColorDisabled=   13748165
      BorderColor     =   8421504
      BorderDrawEvent =   1
      BorderWidth     =   0
      TransparentColor=   14215660
      Caption         =   "X"
      CaptionPosition =   4
      ForeColorDisabled=   8421504
      ForeColorOver   =   192
      ForeColorFocus  =   13003064
      ForeColorDown   =   192
      PictureAlignment=   4
      GradientType    =   2
      TextFadeToColor =   8388608
   End
   Begin MyCommandButton.MyButton btnOk 
      Height          =   495
      Left            =   1935
      TabIndex        =   1
      Top             =   1185
      Visible         =   0   'False
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   873
      BackColor       =   13748165
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   9
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   6805503
      BackColorDisabled=   13748165
      BorderColor     =   32768
      BorderDrawEvent =   1
      BorderWidth     =   0
      TransparentColor=   14215660
      Caption         =   "&Ok"
      CaptionPosition =   4
      ForeColorDisabled=   -2147483632
      ForeColorOver   =   13003064
      ForeColorFocus  =   13003064
      ForeColorDown   =   13003064
      PictureAlignment=   4
      GradientType    =   3
      TextFadeToColor =   8388608
      TextFadeEvents  =   6
   End
   Begin MyCommandButton.MyButton btnNow 
      Height          =   495
      Left            =   1185
      TabIndex        =   2
      Top             =   1185
      Visible         =   0   'False
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   873
      BackColor       =   13748165
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   9
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   13748165
      BackColorDisabled=   13748165
      BorderColor     =   32768
      BorderDrawEvent =   1
      BorderWidth     =   0
      TransparentColor=   14215660
      Caption         =   "&Now"
      CaptionPosition =   4
      ForeColorDisabled=   8421504
      ForeColorOver   =   13003064
      ForeColorFocus  =   13003064
      ForeColorDown   =   13003064
      PictureAlignment=   4
      GradientType    =   3
      TextFadeToColor =   8388608
      TextFadeEvents  =   6
   End
   Begin MyCommandButton.MyButton btnLater 
      Height          =   495
      Left            =   2340
      TabIndex        =   3
      Top             =   1185
      Visible         =   0   'False
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   873
      BackColor       =   13748165
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   9
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   13748165
      BackColorDisabled=   13748165
      BorderColor     =   32768
      BorderDrawEvent =   1
      BorderWidth     =   0
      TransparentColor=   14215660
      Caption         =   "&Later"
      CaptionPosition =   4
      ForeColorDisabled=   8421504
      ForeColorOver   =   13003064
      ForeColorFocus  =   13003064
      ForeColorDown   =   13003064
      PictureAlignment=   4
      GradientType    =   3
      TextFadeToColor =   8388608
      TextFadeEvents  =   6
   End
   Begin VB.Label labelMsg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Checking for Update Now..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   780
      TabIndex        =   4
      Top             =   690
      Width           =   2925
   End
End
Attribute VB_Name = "frmAutoUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public bUpdateAvailable As Boolean
Public sNewRev As String

Private Sub btnClose_Click()
    btnOk_Click

End Sub

Private Sub btnLater_Click()

    Unload Me

End Sub

Private Sub btnNow_Click()
Dim sDBLoc As String
Dim sCopyProg As String
Dim objFile

    Const FILE_ATTRIBUTE_ARCHIVE = &H20
    Const FTP_TRANSFER_TYPE_UNKNOWN = &H0

    Set objFile = CreateObject("Scripting.FileSystemObject")
    
    sDBLoc = "" & GetSetting(App.Title, "DB", App.Title & "DB") & ""
    ' File Location
    
    ' DOWNLOAD UPDATE FILE

    labelMsg.Caption = "Downloading Update Now. Please wait.."
    Screen.MousePointer = 11
    ' SHOW MESSAGE REGARDING DOWNLOAD

    gbOk = FtpGetFile(hConn, "StockWatch Ver" & sNewRev & ".exe", sDBLoc & "\StockWatch Ver" & sNewRev & ".exe", False, FILE_ATTRIBUTE_ARCHIVE, _
        FTP_TRANSFER_TYPE_UNKNOWN, 0&)
        
    Pause 2000
    
    If Not gbOk Then
    ' File may or may not have copied down.
    ' First check to see if the name is present...
    ' then check to see that the length isnt zero...
    ' if its not carry on otherwise try again...
        
        If objFile.FileExists(sDBLoc & "\StockWatch Ver" & sNewRev & ".exe") Then
        
            If FileLen(sDBLoc & "\StockWatch Ver" & sNewRev & ".exe") = 0 Then
        
                gbOk = FtpGetFile(hConn, "StockWatch Ver" & sNewRev & ".exe", sDBLoc & "\StockWatch Ver" & sNewRev & ".exe", False, FILE_ATTRIBUTE_ARCHIVE, _
                    FTP_TRANSFER_TYPE_UNKNOWN, 0&)
            End If
        
        End If
    End If
    
    If gbOk Then
    
        SWdb.Close  ' Close database
        
        Close (lAUDf)   ' close log audit file
        
        sCopyProg = sDBLoc & "/CopyUpdate.exe " & sNewRev
        
        ' SPAWN COPY PROGRAM
        
        Shell sCopyProg, vbNormalNoFocus
    
        Screen.MousePointer = 0
        
        Unload Me
        
        End
    
    Else
        MsgBox "Internet not available. try later for update", vbInformation, "Stock Watch"
        Screen.MousePointer = 0
        
    End If
    
    
End Sub

Private Sub btnOk_Click()

    Unload Me

End Sub

Private Sub Form_Activate()
        
    If CheckForUpdate(sNewRev, False) Then

         btnNow.Visible = True
         btnLater.Visible = True
         btnOk.Visible = False
         labelMsg.Caption = "Update Available. Install?"

    Else
         btnNow.Visible = False
         btnLater.Visible = False
         btnOk.Visible = True
         labelMsg.Caption = "Latest Update Installed"
    End If
    
    Screen.MousePointer = 0

End Sub

Private Sub Form_Load()


    Screen.MousePointer = 11
    ' show busy
    
    
End Sub
