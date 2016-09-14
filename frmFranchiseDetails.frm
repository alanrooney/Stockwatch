VERSION 5.00
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFranDetails 
   BorderStyle     =   0  'None
   Caption         =   "Stockwatch License Manager"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   Icon            =   "frmFranchiseDetails.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFranchiseDetails.frx":1CCA
   ScaleHeight     =   7170
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView Cal 
      Height          =   2820
      Left            =   3285
      TabIndex        =   13
      Top             =   2025
      Visible         =   0   'False
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   4974
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   5606001
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   16777215
      StartOfWeek     =   16842754
      TitleBackColor  =   13342061
      TitleForeColor  =   16777215
      CurrentDate     =   39972
   End
   Begin VB.CheckBox chkSetExpiry 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Set License Expiry Date"
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
      Left            =   3210
      TabIndex        =   28
      Top             =   4350
      Width           =   2460
   End
   Begin VB.CheckBox chkTerminate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Terminate License"
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
      Left            =   405
      TabIndex        =   27
      Top             =   4350
      Width           =   2070
   End
   Begin VB.TextBox txtWarn 
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   5910
      TabIndex        =   25
      ToolTipText     =   "Warn User that License will expire In this Number of Days (or less)"
      Top             =   5535
      Width           =   690
   End
   Begin VB.TextBox txtDays 
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   4005
      TabIndex        =   23
      ToolTipText     =   "License will Auto Expire if Invoices not received within this number of Days"
      Top             =   5550
      Width           =   690
   End
   Begin MSMask.MaskEdBox tedDate 
      Height          =   360
      Left            =   5310
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4875
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MyCommandButton.MyButton btnCal 
      Height          =   360
      Left            =   6270
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4875
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   635
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmFranchiseDetails.frx":87AF
      BackColorDown   =   15133676
      TransparentColor=   14215660
      Caption         =   ""
      DepthEvent      =   1
      PictureDisabled =   "frmFranchiseDetails.frx":BBB4
      PictureDisabledEffect=   2
      ShowFocus       =   -1  'True
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   1845
      TabIndex        =   9
      Top             =   3585
      Width           =   4725
   End
   Begin VB.TextBox txtRegion 
      BackColor       =   &H00F4EFE1&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5865
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3075
      Width           =   690
   End
   Begin VB.TextBox txtPhone 
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   1845
      TabIndex        =   5
      Top             =   3075
      Width           =   2640
   End
   Begin VB.TextBox txtAddress 
      BackColor       =   &H00FFFFFF&
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
      Height          =   1260
      Left            =   1845
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1695
      Width           =   4710
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   1845
      TabIndex        =   1
      Top             =   1215
      Width           =   4710
   End
   Begin MyCommandButton.MyButton btnQuit 
      Height          =   495
      Left            =   4050
      TabIndex        =   11
      Top             =   6375
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
   Begin MyCommandButton.MyButton btnSave 
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      Top             =   6375
      Width           =   2550
      _ExtentX        =   4498
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
      Caption         =   "Save && Send New License"
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
   Begin MyCommandButton.MyButton btnClose 
      Height          =   255
      Left            =   7575
      TabIndex        =   12
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
   Begin MyCommandButton.MyButton btnTerminate 
      Height          =   780
      Left            =   765
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4875
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   1376
      BackColor       =   13748165
      Enabled         =   0   'False
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
      AppearanceMode  =   1
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   13748165
      BackColorDisabled=   13748165
      BorderColor     =   32768
      BorderDrawEvent =   1
      BorderWidth     =   0
      TransparentColor=   14215660
      Caption         =   "Terminate License"
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
   Begin MyCommandButton.MyButton btnSetExpire 
      Height          =   360
      Left            =   3360
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4875
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   635
      BackColor       =   13748165
      Enabled         =   0   'False
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
      AppearanceMode  =   1
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   13748165
      BackColorDisabled=   13748165
      BorderColor     =   32768
      BorderDrawEvent =   1
      BorderWidth     =   0
      TransparentColor=   14215660
      Caption         =   "Set Expiry Date"
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
   Begin VB.Label lblGlow 
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1815
      TabIndex        =   29
      Top             =   1185
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Label lblWarn 
      BackStyle       =   0  'Transparent
      Caption         =   "Warn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5370
      TabIndex        =   26
      Top             =   5580
      Width           =   1935
   End
   Begin VB.Label lblDays 
      BackStyle       =   0  'Transparent
      Caption         =   "Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3480
      TabIndex        =   24
      Top             =   5580
      Width           =   1935
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Watch - Franchise Maintenance Program"
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
      TabIndex        =   22
      Top             =   105
      Width           =   4185
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Franchisee Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   555
      TabIndex        =   21
      Top             =   735
      Width           =   1725
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Joined"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   3735
      TabIndex        =   20
      Top             =   750
      Width           =   1620
   End
   Begin VB.Label lblDateJoined 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   6495
      TabIndex        =   19
      Top             =   750
      Width           =   45
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No License Date Set"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   810
      Left            =   225
      TabIndex        =   18
      Top             =   6255
      Width           =   3660
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblEmail 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   8
      Top             =   3645
      Width           =   1620
   End
   Begin VB.Label lblRegion 
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5175
      TabIndex        =   6
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblPhone 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1215
      TabIndex        =   4
      Top             =   3135
      Width           =   735
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1035
      TabIndex        =   2
      Top             =   1725
      Width           =   915
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1245
      TabIndex        =   0
      Top             =   1275
      Width           =   735
   End
End
Attribute VB_Name = "frmFranDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lFranID As Long
Public sOLDName As String
Public sOLDAddress As String
Public sOLDPhone As String
Public sOLDEmail As String
Public iOLDDays As Integer
Public iOLDWarn As Integer


Private Sub btnCal_Click()
    If Cal.Visible = True Then
        Cal.Visible = False
    Else
        If IsDate(tedDate) Then
            Cal.Value = tedDate
        Else
            Cal.Value = Format(Now, "dd/mm/yy")
        End If
        Cal.Top = btnCal.Top - Cal.Height - 30
        Cal.Left = btnSetExpire.Left
        Cal.Visible = True
    End If

    

End Sub

Private Sub btnClose_Click()

    btnQuit_Click
    

End Sub

Private Sub btnQuit_Click()

    Unload Me

End Sub

Private Sub btnSave_Click()
Dim sNewFileName As String
    
    If FieldsOk() Then
    
        gbOk = SaveLicenseInfo(lFranID)
        
        If MsgBox("Are you sure you want to send this Modified License to " & txtName & " (" & txtRegion & ")", vbYesNo + vbQuestion + vbDefaultButton2, "Franchise License Update") = vbYes Then
        ' confirm transfer
    
            If CreateFranchiseLicenseUpdateFile(sNewFileName, txtRegion) Then
            ' create file based on params above
    
                gbOk = XferFile(sNewFileName, txtRegion)
            End If
        
        End If
    
        Unload Me
    
    End If
    
    Screen.MousePointer = 0
    

End Sub

Private Sub btnSetExpire_Click()
    
    If btnSetExpire.ToggleValue Then
        btnTerminate.ToggleValue = False
    
        btnCal.Enabled = True
            
        btnCal_Click
        
        bSetFocus Me, "btnCal"
        
        showmsg
    
        btnSave.Enabled = True
    
    Else
        
        tedDate.Enabled = False
        btnCal.Enabled = False
    
    End If

End Sub

Private Sub btnTerminate_Click()
    
    Cal.Visible = False
    
    If btnTerminate.ToggleValue Then
        
        btnSetExpire.ToggleValue = False
        tedDate = Format(Now - 1, "dd/mm/yy")

        btnSave.Enabled = True
    
    Else
    
    
    End If

    showmsg

End Sub

Private Sub Cal_DateClick(ByVal DateClicked As Date)
    tedDate.Text = Format(Cal.Value, "dd/mm/yy")
    
    tedDate.Enabled = True
    
    showmsg
    
    Cal.Visible = False

    bSetFocus Me, "btnSave"
    
End Sub


Private Sub Cal_LostFocus()

    Cal.Visible = False

End Sub

Private Sub chkSetExpiry_Click()

    If chkSetExpiry.Value = 1 Then
        chkTerminate.Value = 0
    
        btnTerminate.Enabled = False
    
        btnSetExpire.Enabled = True
    
        tedDate.Enabled = False
        
        btnCal.Enabled = False

        btnTerminate.ToggleValue = False
        btnSetExpire.ToggleValue = False
    
        lblMsg.Caption = "No License Expiry Set"
    
        btnSave.Enabled = False
        
        Cal.Visible = False
    
        txtDays.Enabled = True
        txtWarn.Enabled = True
    
    Else
    
        btnTerminate.Enabled = False
    
        btnSetExpire.Enabled = False
    
        tedDate.Enabled = False
        
        btnCal.Enabled = False

        btnTerminate.ToggleValue = False
        btnSetExpire.ToggleValue = False
    
        lblMsg.Caption = "No License Expiry Set"
    
        btnSave.Enabled = False
        
        Cal.Visible = False
        
        txtDays.Enabled = False
        txtWarn.Enabled = False


    End If
    
    
End Sub




Private Sub Form_Activate()

    btnSave.Enabled = False

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then Unload Me
    

End Sub

Private Sub Form_Load()

    If lFranID <> 0 Then
    
        init
        
        gbOk = ShowFranDetails(lFranID)
        
    End If

End Sub

Public Function ShowFranDetails(lFranID)
Dim rs As Recordset

    Set rs = swDB.OpenRecordset("tblFranchisees")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lFranID
    
    If Not rs.NoMatch Then
    
        txtRegion.Text = rs("Region") & ""
        txtName.Text = rs("Name") & ""
        txtAddress.Text = rs("address") & ""
        txtPhone.Text = rs("Phone") & ""
        txtEmail.Text = rs("Email") & ""
        lblDateJoined.Caption = Format(rs("joined"), "dd mmm yyyy")
    
        sOLDName = rs("Name") & ""
        sOLDAddress = rs("address") & ""
        sOLDPhone = rs("Phone") & ""
        sOLDEmail = rs("Email") & ""
        iOLDDays = rs("Days") & ""
        iOLDWarn = rs("Warn") & ""
    
    End If
    
    rs.Close

Leave:
    Exit Function

ErrorHandler:
    
    MsgBox "Error: " & Trim$(Error)
    Resume Leave
    Resume 0
    
End Function

Public Sub showmsg()

    If btnTerminate.ToggleValue Then
        lblMsg = "License will terminate the next time Franchisee " & txtName & " (" & txtRegion & ") runs Stockwatch and the laptop is online"
    ElseIf btnSetExpire.ToggleValue And IsDate(tedDate) Then
        lblMsg = "Setting New Expiry date " & tedDate & " for Franchisee " & txtName & " (" & txtRegion & ")"
    End If

End Sub

Public Function CreateFranchiseLicenseUpdateFile(sNewFileName As String, sRegion As String)
Dim filenum As Long
Dim sExpiry As String
Dim sFranchiseDetails As String
Dim sEncrypt As String

    On Error GoTo ErrorHandler
    
    ' ENCRYPT LICENSE CONTROL FILE
    
    ' Here the license is encrypted and sent to the franchisee
    ' Any of the items lised below can be modified on the franchisee laptop
    
    ' Note: Region is not modifyable
    
    sNewFileName = sDBLoc & "\" & sRegion & "_Maint.csv"
    
    ' FRANCHISE LICENSE DETAILS
    sFranchiseDetails = "<Name>" & Trim$(txtName.Text) & "/<Name>" & _
            "<Address>" & Trim$(Replace(txtAddress.Text, vbCrLf, "@@")) & "/<Address>" & _
            "<Phone>" & Trim$(txtPhone.Text) & "/<Phone>" & _
            "<Email>" & Trim$(txtEmail.Text) & "/<Email>" & _
            "<Expiry>" & Trim$(tedDate.Text) & "/<Expiry>" & _
            "<Days>" & Trim$(txtDays.Text) & "/<Days>" & _
            "<Warn>" & Trim$(txtWarn.Text) & "/<Warn>"

    ' ENCRYPTED LICENSE
    sEncrypt = Encrypt(sFranchiseDetails, sKey)


    ' CREATE THE FILE
    filenum = FreeFile
    Open sNewFileName For Output As #filenum
    
    Print #filenum, sEncrypt
    Close #filenum
    
    CreateFranchiseLicenseUpdateFile = True
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error " & Trim$(Error)

End Function


Public Function SaveLicenseInfo(lFranID As Long)
Dim rs As Recordset
Dim sExpiry As String

    ' SAVE LICENSE INFO LOCALLY IN SWFRANCHISE SO
    ' ITS NOT ENCRYPTED HERE
    
        
        Set rs = swDB.OpenRecordset("tblFranchisees")
        rs.Index = "PrimaryKey"
        rs.Seek "=", lFranID
        If Not rs.EOF Then
            rs.Edit
        Else
            rs.AddNew
        End If
        
        rs("name") = txtName.Text
        rs("address") = txtAddress.Text
        rs("phone") = txtPhone.Text
        rs("Email") = txtEmail.Text
        rs("Expiry") = tedDate.Text
        rs("Days") = Val(txtDays.Text)
        rs("Warn") = Val(txtWarn.Text)
        
        rs.Update
        rs.Close

    Exit Function
    
ErrorHandler:
    MsgBox "Error " & Trim$(Error)

End Function
Public Function XferFile(sNewFileName As String, sRegion As String)
Dim sPath As String
Dim iCnt As Integer
Dim iMem As Integer
Dim varBkgs As Variant
Dim sBkgs As String
Dim sXferFile As String
Dim sRecvFile As String
Dim iCrLf As Integer
Dim dtBkgDate As Date
Dim bRetry As Boolean
Dim iLogRow As Integer
Dim sUrl As String
Dim sUsername As String
Dim sPassword As String
Dim iTimeout As Integer
Dim sFranchiseDBox As String
Dim objFile As Object

    Screen.MousePointer = 11
    
    On Error GoTo ErrorHandler
    
    Set objFile = CreateObject("Scripting.FileSystemObject")
    ' check for existing Maint file and replace it

    If objFile.FileExists(gbDropBox & "\" & sRegion & "\" & sRegion & "_Maint.csv") Then
    ' see if file already exists
        
        Kill gbDropBox & "\" & sRegion & "_Maint.csv"
    End If
    
    If CopyFileToDropBox(sNewFileName, gbDropBox & "\" & sRegion & "\" & sRegion & "_Maint.csv") Then
                
        MsgBox "File Sent to dropbox"
    
    End If

    Exit Function
    
ErrorHandler:
    MsgBox "Error " & Trim$(Error) & " File not sent to DropBox"
    

End Function


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MoveX = X
    MoveY = Y
    
    SetTranslucent Me.hWnd, 200
    
    bAllowMove = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bAllowMove Then
        Me.Move Me.Left + (X - MoveX), Me.Top + (Y - MoveY)
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bAllowMove = False
    
    SetTranslucent Me.hWnd, 255

End Sub

Public Sub init()
        
    txtRegion.Text = ""
    txtName.Text = ""
    txtAddress.Text = ""
    txtPhone.Text = ""
    txtEmail.Text = ""
    lblDateJoined.Caption = ""
    chkTerminate.Value = 0
    chkSetExpiry.Value = 0
    btnTerminate.Enabled = False
    btnTerminate.ToggleValue = False
    btnSetExpire.Enabled = False
    btnSetExpire.ToggleValue = False
    tedDate = Format(DateAdd("d", 30, Now), "dd/mm/yy")
    tedDate.Enabled = False
    btnCal.Enabled = False
    txtDays.Text = "30"
    txtWarn.Text = "7"
    
End Sub

Private Sub Form_Terminate()
    lFranID = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

    lFranID = 0

End Sub

Private Sub chkTerminate_Click()

    
    If chkTerminate.Value = 1 Then
        
        btnSetExpire.ToggleValue = False
        
        chkSetExpiry.Value = 0
        
        btnTerminate.Enabled = True
    
        btnSetExpire.Enabled = False
    
        tedDate.Enabled = False
        
        btnCal.Enabled = False

        lblMsg.Caption = "No License Expiry Set"
    
        btnSave.Enabled = False
        
        Cal.Visible = False
        
        txtDays.Enabled = False
        txtWarn.Enabled = False

    Else
        btnTerminate.Enabled = False
        btnTerminate.ToggleValue = False

    End If
    


End Sub

Public Function FieldsOk()

    If NameOk() Then
        If AddressOk() Then
            If PhoneOk() Then
                If EmailOk() Then
                    If DaysOk() Then
                        If WarnOk() Then
                            FieldsOk = True
                        Else
                            bSetFocus Me, "txtDays"
                        End If
                    Else
                        bSetFocus Me, "txtDays"
                    End If
                Else
                    bSetFocus Me, "txtEmail"
                End If
            Else
                bSetFocus Me, "txtPhone"
            End If
        Else
            bSetFocus Me, "txtAddress"
        End If
    Else
        bSetFocus Me, "txtName"
    End If
    

End Function

Public Function NameOk()

    If sOLDName <> txtName Then
    
        If MsgBox("Changing Name from " & sOLDName & " To " & txtName & ". Is this Correct?", vbDefaultButton2 + vbYesNo + vbQuestion, "Changing Franchisee Name") = vbYes Then
        
            NameOk = True
        End If
    Else
        NameOk = True

    End If
    
End Function

Public Function AddressOk()

    If sOLDAddress <> txtAddress Then
    
        If MsgBox("Changing Address from " & sOLDAddress & " To " & txtAddress & ". Is this Correct?", vbDefaultButton2 + vbYesNo + vbQuestion, "Changing Franchisee Address") = vbYes Then
        
            AddressOk = True
        End If
    Else
        AddressOk = True

    End If

End Function

Public Function PhoneOk()

    If sOLDPhone <> txtPhone Then
    
        If MsgBox("Changing Phone from " & sOLDPhone & " To " & txtPhone & ". Is this Correct?", vbDefaultButton2 + vbYesNo + vbQuestion, "Changing Franchisee Phone") = vbYes Then
        
            PhoneOk = True
        End If
    Else
        PhoneOk = True
    End If

End Function
Public Function EmailOk()

    If sOLDEmail <> txtEmail Then
    
        If MsgBox("Changing Email from " & sOLDEmail & " To " & txtEmail & ". Is this Correct?", vbDefaultButton2 + vbYesNo + vbQuestion, "Changing Franchisee Email") = vbYes Then
        
            EmailOk = True
        End If
    Else
        EmailOk = True

    End If

End Function
Public Function DaysOk()

    If iOLDDays <> Val(txtDays) Then
    
      If Val(txtDays) > 1 Then
        If MsgBox("Changing Days from " & Trim$(iOLDDays) & " To " & txtDays & ". Is this Correct?", vbDefaultButton2 + vbYesNo + vbQuestion, "Changing Franchisee Days till Expiry") = vbYes Then
        
            DaysOk = True
        End If
      End If
            
    Else
        DaysOk = True
    End If

End Function
Public Function WarnOk()

    If iOLDWarn <> Val(txtWarn) Then
      If Val(txtWarn) > 1 Then
        If MsgBox("Changing Warn from " & Trim$(iOLDWarn) & " To " & txtWarn & ". Is this Correct?", vbDefaultButton2 + vbYesNo + vbQuestion, "Changing Franchisee License Expiry Warn Days") = vbYes Then
        
            WarnOk = True
        End If
      End If
      
    Else
        WarnOk = True
    End If

End Function

Private Sub tedDate_GotFocus()
    gbOk = bSetupControl(Me)

End Sub


Private Sub txtAddress_Change()
    btnSave.Enabled = True

End Sub

Private Sub txtAddress_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    KeyAscii = CharOk(KeyAscii, 2, " ") ' 0 = no only, 1 = char only, 2 = both

End Sub

Private Sub txtAddress_LostFocus()
    lblAddress.ForeColor = sBlack

End Sub

Private Sub txtDays_Change()
    btnSave.Enabled = True

End Sub

Private Sub txtDays_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtDays_KeyPress(KeyAscii As Integer)
    KeyAscii = CharOk(KeyAscii, 0, "") ' 0 = no only, 1 = char only, 2 = both

End Sub

Private Sub txtDays_LostFocus()
    lblDays.ForeColor = sBlack

End Sub

Private Sub txtEmail_Change()
    btnSave.Enabled = True

End Sub

Private Sub txtEmail_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    KeyAscii = CharOk(KeyAscii, 2, "") ' 0 = no only, 1 = char only, 2 = both

End Sub

Private Sub txtEmail_LostFocus()
    lblEmail.ForeColor = sBlack

End Sub

Private Sub txtName_Change()


    btnSave.Enabled = True
End Sub

Private Sub txtName_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

    KeyAscii = CharOk(KeyAscii, 2, " ") ' 0 = no only, 1 = char only, 2 = both
        
End Sub

Private Sub txtName_LostFocus()
    lblName.ForeColor = sBlack

End Sub

Private Sub txtPhone_Change()
    btnSave.Enabled = True

End Sub

Private Sub txtPhone_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
    KeyAscii = CharOk(KeyAscii, 2, " ") ' 0 = no only, 1 = char only, 2 = both

End Sub

Private Sub txtPhone_LostFocus()
    lblPhone.ForeColor = sBlack

End Sub

Private Sub txtWarn_Change()
    btnSave.Enabled = True

End Sub

Private Sub txtWarn_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtWarn_KeyPress(KeyAscii As Integer)
    KeyAscii = CharOk(KeyAscii, 0, "") ' 0 = no only, 1 = char only, 2 = both

End Sub

Private Sub txtWarn_LostFocus()
    lblWarn.ForeColor = sBlack

End Sub
