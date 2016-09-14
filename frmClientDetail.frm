VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmClientDetail 
   BackColor       =   &H00CEEDF4&
   BorderStyle     =   0  'None
   Caption         =   "StockWatch Client Detail"
   ClientHeight    =   7530
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   7980
   Icon            =   "frmClientDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmClientDetail.frx":1CCA
   ScaleHeight     =   7530
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMultipleBars 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D6F5F5&
      Caption         =   "Multiple &Bars"
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
      Height          =   240
      Left            =   4890
      TabIndex        =   9
      ToolTipText     =   "Set if Client is using dual pricing"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CheckBox chkDualPricing 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D6F5F5&
      Caption         =   "&Dual Pricing"
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
      Height          =   240
      Left            =   4935
      TabIndex        =   8
      ToolTipText     =   "Set if Client is using dual pricing"
      Top             =   2730
      Width           =   1410
   End
   Begin VB.TextBox txtMobile 
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
      Height          =   330
      Left            =   5295
      MaxLength       =   20
      TabIndex        =   7
      ToolTipText     =   "Client's Work Number"
      Top             =   2205
      Width           =   2160
   End
   Begin VB.CheckBox chkActive 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D6F5F5&
      Caption         =   "Acti&ve"
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
      Left            =   6510
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1065
      Width           =   915
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
      Height          =   345
      Left            =   1170
      MaxLength       =   100
      TabIndex        =   13
      ToolTipText     =   "Client's Email Address"
      Top             =   4260
      Width           =   6300
   End
   Begin VB.TextBox txtContact 
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
      Height          =   360
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   15
      ToolTipText     =   "Client's Fax Number"
      Top             =   3720
      Width           =   3060
   End
   Begin VB.TextBox txtNotes 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      ToolTipText     =   "Client Notes and till information"
      Top             =   4800
      Width           =   6270
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1185
      MaxLength       =   100
      TabIndex        =   1
      ToolTipText     =   "Client or Company Name"
      Top             =   1065
      Width           =   4935
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
      Height          =   330
      Left            =   5280
      MaxLength       =   20
      TabIndex        =   5
      ToolTipText     =   "Client's Work Number"
      Top             =   1725
      Width           =   2190
   End
   Begin RichTextLib.RichTextBox rtfAddress 
      Height          =   1815
      Left            =   1185
      TabIndex        =   3
      ToolTipText     =   "Client's Address"
      Top             =   1725
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   3201
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      MaxLength       =   200
      Appearance      =   0
      TextRTF         =   $"frmClientDetail.frx":810E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MyCommandButton.MyButton cmdOk 
      Height          =   495
      Left            =   6660
      TabIndex        =   18
      Top             =   6810
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
      TransparentColor=   13561332
      Caption         =   "&Ok"
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
   Begin MyCommandButton.MyButton cmdQuit 
      Height          =   495
      Left            =   5670
      TabIndex        =   19
      Top             =   6810
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
      TransparentColor=   13561332
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
   Begin MyCommandButton.MyButton btnClose 
      Height          =   255
      Left            =   7560
      TabIndex        =   21
      Top             =   120
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
      TransparentColor=   13561332
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
   Begin MSMask.MaskEdBox tedFee 
      Height          =   345
      Left            =   6135
      TabIndex        =   11
      ToolTipText     =   "Fee normally charged to Client"
      Top             =   3720
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   609
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MyCommandButton.MyButton btnBars 
      Height          =   375
      Left            =   6615
      TabIndex        =   23
      Top             =   3165
      Visible         =   0   'False
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   661
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
      TransparentColor=   13561332
      Caption         =   "Bars >"
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
   Begin VB.Label lblFee 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Regular Fee €"
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
      Height          =   270
      Left            =   4755
      TabIndex        =   10
      Top             =   3765
      Width           =   1275
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
      Left            =   1170
      TabIndex        =   22
      Top             =   1005
      Visible         =   0   'False
      Width           =   4980
   End
   Begin VB.Label lblMobile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Mobile"
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
      Height          =   240
      Left            =   4590
      TabIndex        =   6
      Top             =   2235
      Width           =   615
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Email"
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
      Height          =   270
      Left            =   585
      TabIndex        =   12
      Top             =   4290
      Width           =   510
   End
   Begin VB.Label lblContact 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Contact"
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
      Height          =   270
      Left            =   450
      TabIndex        =   14
      Top             =   3765
      Width           =   675
   End
   Begin VB.Label lblNotes 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Till Info and Notes"
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
      Height          =   720
      Left            =   540
      TabIndex        =   16
      Top             =   4815
      Width           =   675
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Name"
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
      Height          =   240
      Left            =   540
      TabIndex        =   0
      Top             =   1110
      Width           =   555
   End
   Begin VB.Label lblAddress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Address"
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
      Height          =   240
      Left            =   360
      TabIndex        =   2
      Top             =   1725
      Width           =   765
   End
   Begin VB.Label lblPhone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Phone"
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
      Height          =   240
      Left            =   4605
      TabIndex        =   4
      Top             =   1710
      Width           =   585
   End
End
Attribute VB_Name = "frmClientDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bInProgress As Boolean
Public lClientID As Long
Public sOldName As String


Private Sub btnBars_Click()

    ' Ver 440 new feature
    

    frmBars.lClientID = lClientID
    frmBars.Show vbModal
    
    bSetFocus Me, "cmdOk"


End Sub

Private Sub btnClose_Click()

    cmdQuit_Click

End Sub

Private Sub chkDualPricing_Click()
        SetInProgress 0, bInProgress

End Sub

Private Sub chkMultipleBars_Click()
        
    ' Ver 440 new
    
        
    SetInProgress 0, bInProgress

    btnBars.Visible = chkMultipleBars.Value

End Sub

Private Sub cmdOk_Click()
Dim bCarryOn As Boolean
Dim objfile As Object

    txtName.Text = Trim$(txtName)
    ' make sure no spaces have ben added front or back of the name

    If FieldsCheckOut() Then
        
        bCarryOn = True
        ' default
        
        If lClientID = -1 Then lClientID = 0
        ' careful with lCustID
        ' must be passed as 0 if its -1 since WriteDB
        ' treats a -1 as if there is only one record in the table
        
        If (sOldName <> txtName) And (sOldName <> "") Then
        
            If MsgBox("Do you wish to Rename Client Name " & sOldName & " As " & txtName, vbYesNo + vbQuestion + vbDefaultButton1, "Rename Client") = vbYes Then
            
                Set objfile = CreateObject("Scripting.FileSystemObject")
                
                If objfile.FolderExists(sDBLoc & "\" & Replace(sOldName, " ", "_")) Then
                ' first make sure the old one existed
           
                    If Not objfile.FolderExists(sDBLoc & "\" & Replace(txtName, " ", "_")) Then
                    ' then make sure the old one does not exist
                
                        Name sDBLoc & "\" & Replace(sOldName, " ", "_") As sDBLoc & "\" & Replace(txtName, " ", "_")
                        ' do rename
                    Else
                        If MsgBox("A Folder for " & Trim$(txtName) & " Already Exists, Continue with Rename?", vbYesNo + vbExclamation + vbDefaultButton1, "New Folder Name Already Exists") = vbNo Then
                            bCarryOn = False
                            ' dont save new name
                            
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If bCarryOn Then
        
        gbOk = WriteDB(Me, "Clients", lClientID, False, 11, _
                    chkActive, _
                    txtName, _
                    rtfAddress, _
                    txtPhone, _
                    txtEmail, _
                    txtContact, _
                    txtMobile, _
                    txtNotes, _
                    tedFee, _
                    chkDualPricing, _
                    chkMultipleBars)
        
        bDualPrice = chkDualPricing
        bMultipleBars = chkMultipleBars
        ' set globally!
        
        LogMsg frmStockWatch, "Client Added/Modified " & txtName, "Addr:" & Replace(rtfAddress.Text, vbCrLf, " ") & " Ph:" & txtPhone & " Mob:" & txtMobile & " Contact:" & txtContact & " Email:" & txtEmail & " Notes:" & Replace(txtNotes, vbCrLf, " ") & " Act:" & Trim$(chkActive) & " Fee: " & Trim$(tedFee)
        
        bInProgress = False
        cmdQuit_Click
    
        bHourGlass False
    
    End If

End Sub

Private Sub cmdQuit_Click()
    ' check if entry in progress and warn
    If bInProgress Then
        If MsgBox("Quit Entering/Modifying Customer Details", vbQuestion + vbYesNo, "Cancel Enter Customer") = vbYes Then
            Unload Me
        End If
    Else
        Unload Me
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    ' return pressed force focus to next available object in tabbing order
        gbOk = GotoNextControl(Me, 0)
    
    ElseIf KeyAscii = 27 Then
        cmdQuit_Click
        
    End If

End Sub

Private Sub Form_Load()

    gbOk = InitClientDetail()
    ' init panel
    
    ' list vars that can be set externally
    
    If lClientID > 0 Then
        
        gbOk = ReadDB(Me, "Clients", lClientID, 11, _
                    chkActive, _
                    txtName, _
                    rtfAddress, _
                    txtPhone, _
                    txtContact, _
                    txtMobile, _
                    txtNotes, _
                    txtEmail, _
                    tedFee, _
                    chkDualPricing, _
                    chkMultipleBars)

        sOldName = txtName
        ' save it incase we're renaming it
        
        bSetFocus Me, "cmdQuit"
    
        bHourGlass False

    End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
        bHourGlass False

End Sub

Private Sub rtfAddress_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub rtfAddress_LostFocus()
    If Right(rtfAddress.Text, 4) = vbCrLf & vbCrLf Then
        rtfAddress.Text = Left(rtfAddress.Text, Len(rtfAddress.Text) - 4)
        rtfAddress.SelStart = 0
    End If
    lblAddress.ForeColor = sBlack

End Sub

Private Sub tedFee_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub tedFee_KeyPress(KeyAscii As Integer)
    KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both

End Sub

Private Sub tedFee_LostFocus()
    lblFee.ForeColor = sBlack

End Sub

Private Sub txtContact_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtContact_KeyPress(KeyAscii As Integer)
    If Not CharCheck(KeyAscii) Then
        KeyAscii = 0
    
    Else
        If Len(txtContact) = 0 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        ElseIf Asc(Right(txtContact, 1)) = 32 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        SetInProgress KeyAscii, bInProgress
    
    End If
    

End Sub

Private Sub txtContact_LostFocus()
    lblContact.ForeColor = sBlack

End Sub

Private Sub txtEmail_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
    
    KeyAscii = CharOk(KeyAscii, 2, "@.-_;") ' 0 = no only, 1 = char only, 2 = both
    
    SetInProgress KeyAscii, bInProgress

End Sub

Private Sub txtEmail_LostFocus()
    lblEmail.ForeColor = sBlack

End Sub

Private Sub txtFee_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtFee_LostFocus()
    lblFee.ForeColor = sBlack

End Sub

Private Sub txtMobile_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtMobile_KeyPress(KeyAscii As Integer)
    
    If Not KeyAscii = 13 Then
            
        SetInProgress KeyAscii, bInProgress
        KeyAscii = CharOk(KeyAscii, 0, " ") ' 0 = no only, 1 = char only, 2 = both
    
    End If

End Sub

Private Sub txtMobile_LostFocus()
    lblMobile.ForeColor = sBlack

End Sub

Private Sub txtPhone_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
    
    If Not KeyAscii = 13 Then
    
        SetInProgress KeyAscii, bInProgress
    
        KeyAscii = CharOk(KeyAscii, 0, " ") ' 0 = no only, 1 = char only, 2 = both
    
    End If
    
End Sub

Private Sub txtPhone_LostFocus()
    lblPhone.ForeColor = sBlack

End Sub

Private Sub txtName_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    
    If Not CharCheck(KeyAscii) Then
        KeyAscii = 0
    
    Else
        If Len(txtName) = 0 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        ElseIf Asc(Right(txtName, 1)) = 32 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        SetInProgress KeyAscii, bInProgress
    
    End If
    

End Sub

Private Sub txtName_LostFocus()
    lblName.ForeColor = sBlack

End Sub

Private Sub txtNotes_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtNotes_KeyPress(KeyAscii As Integer)
    
    SetInProgress KeyAscii, bInProgress
    
End Sub

Private Sub txtNotes_LostFocus()
    lblNotes.ForeColor = sBlack

End Sub

Public Function InitClientDetail()
    
    txtName.Text = ""
    rtfAddress.Text = ""
    txtPhone.Text = ""
    txtContact.Text = ""
    txtMobile.Text = ""
    txtEmail.Text = ""
    txtNotes.Text = ""
    tedFee.Text = ""
    chkActive.Value = 1

End Function

Private Sub rtfAddress_KeyPress(KeyAscii As Integer)

    If Len(rtfAddress.Text) = 0 Then
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
        
    ElseIf Right(rtfAddress.Text, 4) = vbCrLf & vbCrLf Then
        gbOk = GotoNextControl(Me, rtfAddress.TabIndex + 1)
    ElseIf Right(rtfAddress.Text, 2) = vbCrLf And rtfAddress.SelStart = Len(rtfAddress.Text) Then
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    ElseIf Right(rtfAddress.Text, 1) = " " Then
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
    SetInProgress KeyAscii, bInProgress
    
End Sub

Public Function FieldsCheckOut()
    
    If ClientNameOk(txtName.Text) Then
      If Len(txtName.Text) > 2 Then
        
        If Len(rtfAddress.Text) > 3 Then
                
            If Len(rtfAddress.Text) < 200 Then
            
                rtfAddress = Replace(rtfAddress, Chr$(34), "'")
                txtNotes.Text = Replace(txtNotes, Chr$(34), "'")
                
                If Trim$(txtEmail) = "" Or InStr(txtEmail, "@") > 0 Then
                    
                    FieldsCheckOut = True
                Else
                    MsgBox "Please Enter a Valid Email Address"
                    bSetFocus Me, "txtEmail"
                End If
            
            Else
                MsgBox "Address must be less than 200 Characters"
                bSetFocus Me, "rtfAddress"
            End If
            
        Else
            MsgBox "Must have a minimum of 1 Address line"
            bSetFocus Me, "rtfAddress"
        End If
    
      Else
        MsgBox "Customer Name must be at least 3 Characters long"
        bSetFocus Me, "txtName"
      End If
    Else
      bSetFocus Me, "txtName"
    End If
    
End Function

Public Sub SetInProgress(iKeyAsc As Integer, bInProgress As Boolean)

    If iKeyAsc <> Asc(vbCr) Then
         bInProgress = True
         cmdOk.Enabled = True
    End If
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MoveX = X
    MoveY = Y
    
    SetTranslucent Me.hwnd, 200
    
    bAllowMove = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bAllowMove Then
        Me.Move Me.Left + (X - MoveX), Me.Top + (Y - MoveY)
    End If
    



End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bAllowMove = False
    
    SetTranslucent Me.hwnd, 255

End Sub


Public Function ClientNameOk(sClient As String)
Dim rs As Recordset
    
    On Error GoTo ErrorHandler
    
    ' Make sure there isnt a duplicate name here
    '                                                       LIKE " & """" & sPLU & "*""" & " ORDER
    Set rs = SWdb.OpenRecordset("SELECT ID FROM tblClients WHERE txtName = " & """" & sClient & """")
    If Not (rs.EOF And rs.BOF) Then
    
        If rs("ID") <> lClientID Then
        ' As long as we're not doing an edit of a client
        ' then warn of duplication....
            
            
            MsgBox "This Client Name is already in Use"
        
        Else
            ClientNameOk = True
        End If
        
    Else
        ClientNameOk = True
    End If
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("ClientNameOk") Then Resume 0
    Resume CleanExit





End Function
