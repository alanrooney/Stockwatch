VERSION 5.00
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmFranchiseEmail 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6465
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmFranchiseEmail.frx":0000
   ScaleHeight     =   6480
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   1530
      Top             =   5610
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   705
      Top             =   5625
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.TextBox txtEmailFrom 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   540
      MaxLength       =   100
      TabIndex        =   1
      Top             =   1875
      Width           =   5400
   End
   Begin VB.TextBox txtEmailTo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   540
      MaxLength       =   100
      TabIndex        =   3
      Top             =   2655
      Width           =   5400
   End
   Begin VB.TextBox txtSubj 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   540
      MaxLength       =   100
      TabIndex        =   5
      Text            =   "Audit Report"
      Top             =   3420
      Width           =   5400
   End
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   540
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "frmFranchiseEmail.frx":53F8
      Top             =   4170
      Width           =   5400
   End
   Begin MyCommandButton.MyButton cmdSend 
      Height          =   495
      Left            =   5205
      TabIndex        =   8
      Top             =   5625
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
      BackColorFocus  =   6805503
      BackColorDisabled=   13748165
      BorderColor     =   32768
      BorderDrawEvent =   1
      BorderWidth     =   0
      TransparentColor=   14215660
      Caption         =   "&Send"
      CaptionPosition =   4
      ForeColorDisabled=   -2147483629
      ForeColorOver   =   13003064
      ForeColorFocus  =   13003064
      ForeColorDown   =   13003064
      PictureAlignment=   4
      GradientType    =   3
      TextFadeToColor =   8388608
      TextFadeEvents  =   6
   End
   Begin MyCommandButton.MyButton btnClose 
      Height          =   255
      Left            =   5970
      TabIndex        =   9
      Top             =   90
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
   Begin MyCommandButton.MyButton btnQuit 
      Height          =   495
      Left            =   4200
      TabIndex        =   14
      Top             =   5625
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
      Caption         =   "&Quit"
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
   Begin VB.Label lblReportStatus 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5055
      TabIndex        =   13
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblReportName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1320
      TabIndex        =   12
      Top             =   1125
      Width           =   4605
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   585
      TabIndex        =   11
      Top             =   1200
      Width           =   690
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stockwatch - Franchise Report Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   525
      TabIndex        =   10
      Top             =   105
      Width           =   3270
   End
   Begin VB.Label lblEmailFrom 
      BackStyle       =   0  'Transparent
      Caption         =   "Email &From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   540
      TabIndex        =   0
      Top             =   1620
      Width           =   3030
   End
   Begin VB.Label lblEmailTo 
      BackStyle       =   0  'Transparent
      Caption         =   "&Email To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   540
      TabIndex        =   2
      Top             =   2415
      Width           =   3030
   End
   Begin VB.Label lblSubj 
      BackStyle       =   0  'Transparent
      Caption         =   "Su&bject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   540
      TabIndex        =   4
      Top             =   3150
      Width           =   3030
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "&Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   540
      TabIndex        =   6
      Top             =   3900
      Width           =   3030
   End
End
Attribute VB_Name = "frmFranchiseEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()

    Unload Me

End Sub

Private Sub btnQuit_Click()

    Unload Me

End Sub

Private Sub cmdSend_Click()
    
    
 On Error GoTo ErrorHandler

 lblReportStatus.Caption = "Sending..."
 
 SaveSetting appname:="Stockwatch", Section:="Email", Key:="FromEmail", Setting:=txtEmailFrom.Text
 SaveSetting appname:="Stockwatch", Section:="Email", Key:="AuditEmail", Setting:=txtEmailTo.Text
 
 'gbOk = SendEMail(txtEmailTo, txtEmailFrom, txtSubj, txtMessage, sDBLoc & "\AuditReport.Doc", False)
 
 gbOk = SendMail(txtEmailTo, txtSubj, txtMessage, sDBLoc & "\AuditReport.Doc")
 
 MsgBox "Email is Queued. Open Outlook and press Send/Receive to Send Email"
 
 bSetFocus Me, "btnQuit"
 
 Exit Sub
    
ErrorHandler:

    MsgBox "Problem Emailing Report (" & sDBLoc & "\AuditReport.Doc) " & Trim$(Error)


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    ' return pressed force focus to next available object in tabbing order
        gbOk = GotoNextControl(Me, 0)
    
    
    ElseIf KeyAscii = 27 Then
        Unload Me
        
    End If



End Sub

Private Sub Form_Load()

    GetEmailDefaults
    
    txtEmailFrom.Text = gbfromAddress
    txtEmailTo.Text = gbAuditEmail
    lblReportName.Caption = sDBLoc & "\AuditReport.Doc"
    ' get email to from registry settings


End Sub

 Private Function SendMail(sSendTo As String, sSubject As String, sMessage As String, sAttachPath As String)
 'KB113033 How to Send a Mail Message Using Visual Basic MAPI Controls
 'MAPI constants from CONSTANT.TXT file:
 Const ATTACHTYPE_DATA = 0
 Const RECIPTYPE_TO = 1
 Const RECIPTYPE_CC = 2

 On Error GoTo errh

 'Open up a MAPI session:
 Me.MAPISession1.DownLoadMail = False 'revised - sending mail
 Me.MAPISession1.SignOn
 'Point the MAPI messages control to the open MAPI session:
 Me.MAPIMessages1.SessionID = Me.MAPISession1.SessionID

 'Create a new message
 Me.MAPIMessages1.MsgIndex = -1 'revised
 Me.MAPIMessages1.Compose

 'Set the subject of the message:
 Me.MAPIMessages1.MsgSubject = sSubject
 'Set the message content:
 Me.MAPIMessages1.MsgNoteText = sMessage

 'The following four lines of code add an attachment to the message,
 'and set the character position within the MsgNoteText where the
 'attachment icon will appear. A value of 0 means the attachment will
 'replace the first character in the MsgNoteText. You must have at
 'least one character in the MsgNoteText to be able to attach a file.
 Me.MAPIMessages1.AttachmentIndex = 0
 'Set the type of attachment:
 Me.MAPIMessages1.AttachmentType = ATTACHTYPE_DATA
 'Set the icon title of attachment:
' Me.MAPIMessages1.AttachmentName = sAttachName
 'Set the path and file name of the attachment:
 Me.MAPIMessages1.AttachmentPathName = sAttachPath

 'Set the recipients
 Me.MAPIMessages1.RecipIndex = 0
 Me.MAPIMessages1.RecipType = RECIPTYPE_TO
 Me.MAPIMessages1.RecipDisplayName = sSendTo '4/22/03
 'Me.MAPImessages1.RecipAddress = sSendTo

 'MESSAGE_RESOLVENAME checks to ensure the recipient is valid and puts
 'the recipient address in MapiMessages1.RecipAddress
 'If the E-Mail name is not valid, a trappable error will occur.
 'Me.MAPImessages1.ResolveName 'comment out due to receiptent error w/ GW6.5 1/5/04
 'Send the message:
 Me.MAPIMessages1.Send True 'revised

xit:
 'Close MAPI mail session:
 Me.MAPISession1.SignOff
xit2:
 Screen.MousePointer = 0
 Exit Function

errh:
 If Err.Number = 32053 Then Resume xit2
 MsgBox Err.Description, vbCritical, Err.Number
 Resume xit
 End Function

Public Function SendEMail(sEmailTo As String, _
                            sEmailFrom As String, _
                            sSubject As String, _
                            sBody As String, _
                            sFilePath As String, _
                            bMultipleattachments As Boolean) As String
      
    On Error GoTo SendMail_Error:
    
    Dim iRow As Integer
    Dim lobj_cdomsg As CDO.Message
    
    Screen.MousePointer = 11
    
    Set lobj_cdomsg = New CDO.Message
    lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = gbSMTP    ' email setting
    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = gbPort ' email setting
    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = gbSSL     ' email setting
    lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
    If gbUsername <> "" Then
        lobj_cdomsg.Configuration.Fields(cdoSendUserName) = gbUsername ' email setting
    End If
    
    If gbPassword <> "" Then
        lobj_cdomsg.Configuration.Fields(cdoSendPassword) = gbPassword ' email setting
    End If
    
    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 60
    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
    lobj_cdomsg.Configuration.Fields.Update
    lobj_cdomsg.To = sEmailTo   ' passed
    lobj_cdomsg.From = sEmailFrom  ' from License
    lobj_cdomsg.Subject = sSubject   ' passed
    lobj_cdomsg.TextBody = sBody   ' passed
    
    lobj_cdomsg.AddAttachment (sFilePath)    ' passed
        
    lobj_cdomsg.Send
    Set lobj_cdomsg = Nothing
    SendEMail = True

    Screen.MousePointer = 0

leave:
    Exit Function
          
SendMail_Error:
    MsgBox "Error sending email " & Trim$(Error)
    Resume leave
    
End Function


Public Function GetEmailDefaults()

        gbSMTP = GetSetting("Stockwatch", "Email", "smtp")
        gbPort = Val(GetSetting("Stockwatch", "Email", "port"))
        gbfromAddress = GetSetting("Stockwatch", "Email", "FromEmail")

        gbSSL = Val(GetSetting("Stockwatch", "Email", "ssl"))
        gbUsername = GetSetting("Stockwatch", "Email", "username")
        gbPassword = GetSetting("Stockwatch", "Email", "password")
        gbAuditEmail = GetSetting("Stockwatch", "Email", "AuditEmail")
    
End Function

