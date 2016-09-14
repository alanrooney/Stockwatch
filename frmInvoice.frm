VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmInvoice 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   Icon            =   "frmInvoice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmInvoice.frx":1CCA
   ScaleHeight     =   9750
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView Cal 
      Height          =   2820
      Left            =   420
      TabIndex        =   20
      Top             =   5220
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
      StartOfWeek     =   16515074
      TitleBackColor  =   10914457
      TitleForeColor  =   16777215
      CurrentDate     =   39972
   End
   Begin MyCommandButton.MyButton btnCal 
      Height          =   315
      Left            =   2085
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4860
      Visible         =   0   'False
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmInvoice.frx":1136A
      BackColorDown   =   15133676
      TransparentColor=   14215660
      Caption         =   ""
      DepthEvent      =   1
      PictureDisabled =   "frmInvoice.frx":1476F
      PictureDisabledEffect=   2
      ShowFocus       =   -1  'True
   End
   Begin VB.TextBox txtDescription 
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
      Left            =   2430
      MaxLength       =   50
      TabIndex        =   0
      Top             =   4830
      Width           =   4695
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   2  'Center
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
      Left            =   7290
      MaxLength       =   6
      TabIndex        =   1
      Top             =   4830
      Width           =   1260
   End
   Begin MyCommandButton.MyButton btnCreateInvoice 
      Height          =   495
      Left            =   7005
      TabIndex        =   2
      Top             =   9015
      Width           =   1590
      _ExtentX        =   2805
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
      Caption         =   "Create &Invoice"
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
   Begin VB.Label labelInvoiceNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
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
      Left            =   7425
      TabIndex        =   13
      Top             =   3885
      Width           =   1170
   End
   Begin VB.Label labelIBAN 
      BackColor       =   &H00FFFFFF&
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
      Height          =   300
      Left            =   1410
      TabIndex        =   28
      Top             =   8280
      Width           =   3915
   End
   Begin VB.Label labelBIC 
      BackColor       =   &H00FFFFFF&
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
      Height          =   300
      Left            =   1410
      TabIndex        =   27
      Top             =   7950
      Width           =   3915
   End
   Begin VB.Label labelAccountName 
      BackColor       =   &H00FFFFFF&
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
      Height          =   300
      Left            =   1410
      TabIndex        =   26
      Top             =   6885
      Width           =   3915
   End
   Begin VB.Label labelSortCode 
      BackColor       =   &H00FFFFFF&
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
      Height          =   300
      Left            =   1410
      TabIndex        =   25
      Top             =   7590
      Width           =   3915
   End
   Begin VB.Label labelAccountNo 
      BackColor       =   &H00FFFFFF&
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
      Height          =   300
      Left            =   1410
      TabIndex        =   24
      Top             =   7245
      Width           =   3915
   End
   Begin VB.Label labelBank 
      BackColor       =   &H00FFFFFF&
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
      Height          =   300
      Left            =   1410
      TabIndex        =   23
      Top             =   6525
      Width           =   3915
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Left            =   465
      TabIndex        =   22
      Top             =   4845
      Width           =   1575
   End
   Begin VB.Label labelWebSite 
      Alignment       =   2  'Center
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
      Height          =   270
      Left            =   405
      TabIndex        =   19
      Top             =   1875
      Width           =   3795
   End
   Begin VB.Label labelFranEmail 
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
      Height          =   270
      Left            =   420
      TabIndex        =   18
      Top             =   1650
      Width           =   4530
   End
   Begin VB.Label labelVatRate 
      BackStyle       =   0  'Transparent
      Caption         =   "23%"
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
      Left            =   6660
      TabIndex        =   17
      Top             =   6795
      Width           =   450
   End
   Begin VB.Label labelFranAddress 
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
      Height          =   1275
      Left            =   6315
      TabIndex        =   16
      Top             =   1155
      Width           =   2130
   End
   Begin VB.Label labelClientAddress 
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
      Height          =   870
      Left            =   495
      TabIndex        =   15
      Top             =   3225
      Width           =   2970
   End
   Begin VB.Label labelDate 
      BackStyle       =   0  'Transparent
      Caption         =   "10 Jan 2012"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7410
      TabIndex        =   14
      Top             =   3330
      Width           =   1170
   End
   Begin VB.Label labelFranRegion 
      BackStyle       =   0  'Transparent
      Caption         =   "SL1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7425
      TabIndex        =   12
      Top             =   3615
      Width           =   1155
   End
   Begin VB.Label labelInvoiceLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00D1D1E9&
      Caption         =   "INVOICE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4035
      TabIndex        =   11
      Top             =   2310
      Width           =   975
   End
   Begin VB.Label labelFee 
      Alignment       =   1  'Right Justify
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
      Height          =   300
      Left            =   7335
      TabIndex        =   10
      Top             =   6405
      Width           =   1050
   End
   Begin VB.Label labelFranName 
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
      Height          =   270
      Left            =   6315
      TabIndex        =   9
      Top             =   855
      Width           =   2130
   End
   Begin VB.Label labelFranPhone 
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
      Height          =   270
      Left            =   6315
      TabIndex        =   8
      Top             =   2445
      Width           =   2130
   End
   Begin VB.Label labelClientName 
      BackColor       =   &H00FFFFFF&
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
      Height          =   300
      Left            =   495
      TabIndex        =   7
      Top             =   2940
      Width           =   2970
   End
   Begin VB.Label labelVat 
      Alignment       =   1  'Right Justify
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
      Height          =   300
      Left            =   7335
      TabIndex        =   6
      Top             =   6750
      Width           =   1050
   End
   Begin VB.Label labelTotal 
      Alignment       =   1  'Right Justify
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
      Height          =   300
      Left            =   7335
      TabIndex        =   5
      Top             =   7095
      Width           =   1050
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
      Height          =   420
      Left            =   420
      TabIndex        =   4
      Top             =   4815
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Watch - Invoice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   300
      TabIndex        =   3
      Top             =   15
      Width           =   2310
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lDatesID As Long
Public lClientID As Long
Public bNoFeeSet As Boolean

Public Function ShowInvoice(lDatesID As Long, lClientID As Long)
Dim rs As Recordset
Dim sName As String
Dim sAddress As String
Dim sPhone As String
Dim sEmail As String
Dim sCount As String
Dim dtExpiry As Date
Dim sClientAddress As String
Dim iDays As Integer
Dim iwarn As Integer
Dim sBank As String
Dim sNameOnAccount As String
Dim sAccountNo As String
Dim sSortCode As String
Dim sBIC As String
Dim sIBAN As String

    
    ' GET INVOICE IF ITS THERE ALREADY
        
    ' ver 3.0.6 (2) added '&' if there was already one '&' in names
        
    ' ver 3.0.6 (3) sEmail is picked up from franchise license info
    ' so email should be set to region@stockwatchireland.ie
    
    InitInvoice

    labelVatRate.Caption = Trim$(sngvatrate) & "%"
    ' Show current vat rate
    
    ' FRANCHISEE INFORMATION
    
    If GetLicenseInfo(sName, sAddress, sPhone, sEmail, dtExpiry, iDays, iwarn) Then
        labelFranName = Replace(sName, "&", "&&")
        labelFranAddress = sAddress
        labelFranPhone = sPhone
        labelFranRegion = gbRegion
        labelFranEmail = sEmail
    End If
        
    If GetBankInfo(sBank, sNameOnAccount, sAccountNo, sSortCode, sBIC, sIBAN) Then
        labelBank = sBank
        labelAccountName = Replace(sNameOnAccount, "&", "&&")
        labelAccountNo = sAccountNo
        labelSortCode = sSortCode
        labelBIC = sBIC
        labelIBAN = sIBAN
    End If
    
    ' CUSTOMER INFORMATION
    
    labelClientName.Caption = GetClientName(lClientID, sClientAddress, False)
    labelClientName.Caption = Replace(labelClientName.Caption, "&", "&&")
    labelClientAddress.Caption = sClientAddress
    txtAmount = Format(GetClientDefaultFee(lClientID), "0.00")
    ' Customer Name, Address & default Fee (Client Table)
            
    If Val(txtAmount) = 0 Then
        bNoFeeSet = True
    End If
    
    
    lblDate.Caption = GetAuditOnDate(lDatesID)
    ' date audit was carried out
    
    labelDate.Caption = Format(lblDate.Caption, sDMMYY)
    ' Set Date of Invoice = date audit was carried out on
    ' ver 3.0.6 (1)
    
    txtDescription = GetDefaultInvoiceText()
    ' Default Invoice Text (from Invoice Default Tbl)
            
End Function

Private Sub btnCal_Click()
    
    bSetFocus Me, "btnCal"
    
    If Cal.Visible = True Then
        Cal.Visible = False
    Else
        If IsDate(lblDate) Then
            Cal.Value = lblDate
        Else
            Cal.Value = Format(Now, sDMY)
        End If
        Cal.Top = btnCal.Top + btnCal.Height + 30
        Cal.Left = lblDate.Left
        Cal.Visible = True
    End If


End Sub



Private Sub btnCreateInvoice_Click()
Dim sFolder As String
'Dim sInvoiceFileName As String
Dim lInvoiceID As Long
Dim sFranRegion As String
Dim sInvoiceNo As String
Dim sDate As String
Dim sTotalFee As String
Dim sInvoiceSummaryFile As String
Dim sClient As String

    btnCreateInvoice.Enabled = False
    
    Screen.MousePointer = 11
    
    ' CREATE INVOICE - WORD DOCUMENT in CLIENT FOLDER
    ' File Name: SL1-1000inv (area-invoice number)
    
'    frmStockWatch.timSummaryFiles.Enabled = False

    
    If fieldsok() Then
    
    
      LogMsg Me, "Saving Invoice", ""
      
      If SaveInvoice(lInvoiceID, lDatesID) Then
    
        If CreateClientFolder(sFolder, frmStockWatch.lblClient.Tag, Replace(frmStockWatch.labelTo.Tag, "/", "-")) Then
        ' Check is Folder created  - if not create it
        ' i.e. \StockWatch\Client Name\
    
            LogMsg Me, "Creating Invoice", ""
            
            If CreateInvoice(lInvoiceID, sFolder) Then
        
                If bNoFeeSet Then
                    If MsgBox("Do you wish to set €" & Trim$(txtAmount) & " as a default fee for this client?", vbDefaultButton1 + vbYesNo + vbQuestion, "Set Default Fee for Client") = vbYes Then
                        gbOk = SetDefaultFee(lClientID, txtAmount)
'Ver 405                Was using the lSelClientID variable
'                       and may have been causing a problem if switching between clients
'
                    
                    End If
                End If
                
                On Error Resume Next
                
                If Not gbTestEnvironment Then
                    gbOk = RestartAgentProgram()
                
                    ' CHECK that SWIAgent is running
                    ' If not start it

                    gbOk = SendInvoiceBySMTP()
                
                End If
                
                LogMsg Me, "Invoice Created", ""
                
            End If
        End If
      

      End If
    
      Unload Me
    
    End If

leave:
    
'    btnCreateInvoice.Enabled = True
' Ver405 no need for this since we've unloaded the form
    
    Screen.MousePointer = 0
    
    Exit Sub
    

ErrorHandler:
    LogMsg Me, "", Trim$(Error) & " Create Invoice Btn"
    Resume leave
    

End Sub

Private Sub Cal_DateClick(ByVal DateClicked As Date)
    lblDate.Caption = Format(Cal.Value, sDMY)
    
    Cal.Visible = False
    btnCal.Visible = False
    
    bSetFocus Me, "txtDescription"

End Sub

Private Sub Form_Activate()

    If Val(txtAmount) = 0 Then
        bSetFocus Me, "txtAmount"
    Else
        bSetFocus Me, "btnCreateInvoice"
    End If

End Sub

Private Sub Form_Load()

    gbOk = ShowInvoice(lDatesID, lClientID)

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    ' return pressed force focus to next available object in tabbing order
        gbOk = GotoNextControl(Me, 0)
    
    ElseIf KeyAscii = 27 Then
        Unload Me
        
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

Public Function InitInvoice()

        labelClientName.Caption = ""
        labelClientAddress.Caption = ""
        labelFranName.Caption = ""
        labelFranAddress.Caption = ""
        labelFranPhone.Caption = ""
        
        labelFranRegion.Caption = ""
        labelInvoiceNumber.Caption = ""
        labelDate.Caption = ""
        
        
        labelVatRate.Caption = ""
        lblDate.Caption = ""
        txtDescription.Text = ""
        txtAmount.Text = ""
        
        labelFee.Caption = ""
        labelVat.Caption = ""
        labelTotal.Caption = ""
        
        bNoFeeSet = False

End Function

Private Sub labelFee_Change()
    labelVat = Format(Val(labelFee) * Val(labelVatRate) / 100, "0.00")

    labelTotal = Format(Val(labelFee) + Val(labelVat), "0.00")

End Sub

Private Sub lblDate_Click()
    gbOk = bSetupControl(Me)

    btnCal.Visible = True

End Sub

Private Sub txtAmount_Change()
    labelFee = Format(txtAmount, "0.00")

End Sub

Private Sub txtAmount_GotFocus()
    gbOk = bSetupControl(Me)
    btnCal.Visible = False

End Sub


Private Sub txtAmount_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        bSetFocus Me, "btnCreateInvoice"
    
    Else
        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
    End If

End Sub

Private Sub txtDescription_GotFocus()
    gbOk = bSetupControl(Me)
    btnCal.Visible = False
    
End Sub

Public Function SaveInvoice(lInvoiceID As Long, lDtsID As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler

    Set rs = SWdb.OpenRecordset("tblDates")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lDtsID
    If Not rs.NoMatch Then
    
        rs.Edit
        rs("InvNumber") = Val(GetNextInvoiceNumber())
        labelInvoiceNumber.Caption = rs("InvNumber") & ""
        
        rs("INVDate") = labelDate.Caption
        rs("ON") = lblDate.Caption
        rs("INVDescription") = txtDescription
        rs("INVAmount") = Val(txtAmount)
        rs("INVVat") = Val(labelVat)
        rs("INVTotal") = Val(labelTotal)
        rs("INVVatRate") = sngvatrate
        rs("InvNotYetSentToStockWatch") = True  ' for summary file thur internet server
        rs("InvSMTPEmailNotSentYet") = True ' for email
        
        rs.Update
        rs.Bookmark = rs.LastModified
        lInvoiceID = rs("ID") + 0
    
    End If
    
    rs.Close

'    sInvoiceFileName = labelFranRegion.Caption & "_" & labelInvoiceNumber.Caption
    
    SaveInvoice = True

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("SaveInvoice") Then Resume 0
    Resume CleanExit
End Function

Public Function GetNextInvoiceNumber() As Long
Dim rs As Recordset
Dim bAsk As Boolean

    On Error GoTo ErrorHandler

    bAsk = False
    Set rs = SWdb.OpenRecordset("Select InvNumber FROM tblDates ORDER BY INVNumber DESC")
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        If Not IsNull(rs("InvNumber")) Then
            If Val(rs("InvNumber")) = 0 Then
                bAsk = True
            Else
                GetNextInvoiceNumber = rs("InvNumber") + 1
            End If
        Else
            bAsk = True
        End If
    Else
            bAsk = True
    End If
    rs.Close
    
    If bAsk Then
        GetNextInvoiceNumber = InputBox("Set first Invoice Number", "Invoice Number", "100")
    End If
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetNextInvoiceNumber") Then Resume 0
    Resume CleanExit
            
End Function

Public Function CreateInvoice(lID As Long, sClientFolder As String)
Dim rs As Recordset
Dim objfile As Object
Dim Reportdoc
Dim sReport As String
Dim sNextPage As String

Dim iTotPages As Integer
Dim sAddr As String
Dim sClient As String
Dim iRow As Integer
Dim iCol As Integer
Dim sFrom As String
Dim sTo As String
Dim sName As String
Dim sAddress As String
Dim sPhone As String
Dim sEmail As String
Dim dtExpiry As Date
Dim iDays As Integer
Dim iwarn As Integer
Dim sBank As String
Dim sNameOnAccount As String
Dim sAccountNo As String
Dim sSortCode As String
Dim sBIC As String
Dim sIBAN As String

    
    On Error GoTo ErrorHandler
    
    Set objfile = CreateObject("Scripting.FileSystemObject")
    ' create object

    bHourGlass True
    
    frmStockWatch.labelTitle.Caption = "Creating Invoice"
    
'    sTitle = Replace(sTitle, "Print ", "")
    ' incase its a print sheet thats looked for
    
    sReport = sClientFolder & "/Invoice.Doc"
   
    On Error Resume Next
            
    Kill sReport
'    Kill sNextPage
    ' DELETE Report HERE!
        '
    On Error GoTo ErrorHandler
    
    gbOk = TerminateWINWORD()
    
    Set WriteWord = New Word.Application
    
    WriteWord.Visible = False
    
'    iTotPages = 1
    
     bHourGlass True
     
     Set Reportdoc = WriteWord.Documents.Add(sDBLoc & "\Templates\Invoice.dot")
     
     WriteWord.Visible = False
     WriteWord.ActiveDocument.UndoClear
    
     With WriteWord.ActiveDocument.Bookmarks
    
        Set rs = SWdb.OpenRecordset("Select * FROM tblDates INNER JOIN tblClients ON tblDates.ClientID = tblClients.ID WHERE tblDates.ID = " & Trim$(lID))
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
            
            .Item("Client").Range.Text = rs("txtName") & ""
            .Item("Address").Range.Text = rs("rtfAddress") & ""
            
            If GetLicenseInfo(sName, sAddress, sPhone, sEmail, dtExpiry, iDays, iwarn) Then
            
                .Item("FranName").Range.Text = sName
                .Item("FranAddress").Range.Text = sAddress
                .Item("FranPhone").Range.Text = sPhone
                .Item("FranEmail").Range.Text = sEmail
            End If
            
            .Item("Area").Range.Text = gbRegion
            .Item("InvoiceNo").Range.Text = rs("InvNumber") & ""
            .Item("DateOfAudit").Range.Text = Format(rs("on"), sDMMYY)
            .Item("Date").Range.Text = Format(Now, sDMMYY)
            
            .Item("VatRate").Range.Text = Trim$(rs("INVVatRate")) & "%"
            .Item("Description").Range.Text = rs("INVDescription") & ""
            .Item("Amt1").Range.Text = Format(rs("INVAmount"), "Currency")
            .Item("SubTotal").Range.Text = Format(rs("INVAmount"), "Currency")
            .Item("Vat").Range.Text = Format(rs("INVVat"), "Currency")
            .Item("Total").Range.Text = Format(rs("INVTotal"), "Currency")
            
            If GetBankInfo(sBank, sNameOnAccount, sAccountNo, sSortCode, sBIC, sIBAN) Then
                .Item("Bank").Range.Text = sBank
                .Item("NameOnAccount").Range.Text = sNameOnAccount
                .Item("AccountNo").Range.Text = sAccountNo
                .Item("SortCode").Range.Text = sSortCode
                .Item("BIC").Range.Text = sBIC
                .Item("IBAN").Range.Text = sIBAN
            
            End If

'            .Item("Note").Range.Text = rs("Note") & ""
            
            rs.Close
        
            WriteWord.Visible = False
    
            WriteWord.Visible = False
        
            WriteWord.ActiveDocument.UndoClear
    
            WriteWord.Application.NormalTemplate.Saved = True
            
            WriteWord.Application.ActiveDocument.SaveAs (sReport)
            ' Save it as the file name thats required
        
            Reportdoc.Close False
            ' close first page
        End If
    End With
    
    

CloseWordStuff:
    
    bHourGlass True
    
    On Error Resume Next
    
    WriteWord.Application.NormalTemplate.Saved = True
    
    WriteWord.Quit vbTrue
    
    Set WriteWord = Nothing
    
    Set objfile = Nothing
    Set Reportdoc = Nothing
'    Set NextPageDoc = Nothing
    
    
    Kill sNextPage
    
    CreateInvoice = True

CleanExit:
'
    
    bHourGlass True
    
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
'    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("CreateInvoice") Then Resume 0
    Resume CloseWordStuff

End Function


'''Public Function SendMail(sSubj As String, sEmailFrom As String, sEmailTo As String, sBodyText As String, sAttachment As String)
'''Dim iRow As Integer
'''Dim sFileName As String
'''Dim sFiles As String
'''Dim objMessage As Object
'''Dim sBody As String
'''
'''Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
'''Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).
'''Const cdoAnonymous = 0 'Do not authenticate
'''Const cdoBasic = 1 'basic (clear-text) authentication
'''Const cdoNTLM = 2 'NTLM
'''
''''This sample sends a simple HTML EMail Message with Attachment File via GMail servers using SMTP authentication and SSL.
''''It's like any other mail but requires that you set the SMTP Port to 465 and tell CDO to use SSL
''''By Hackoo © 2011
'''
'''    On Error Resume Next
'''
'''    Set objMessage = CreateObject("CDO.Message")
'''    objMessage.Subject = sSubj
''''    objMessage.From = """Me"" <" & sEmailFrom & ">"
'''    objMessage.From = "" & sEmailFrom & " <" & sEmailFrom & ">"
'''    objMessage.To = sEmailTo 'change this To
'''
'''    'objMessage.CC = "dest2@yahoo.fr" 'change this CC means : Carbon Copy
'''    'objMessage.BCC = "dest3@gmail.com" 'change this BCC means : Blink Carbon Copy
'''    sBody = "<center><font size=4 FACE=Tahoma Color=black>" & sBodyText
'''    objMessage.HTMLBody = sBody
'''
'''    objMessage.AddAttachment sAttachment
'''
'''    '==This section provides the configuration information for the remote SMTP server.
'''
'''    objMessage.Configuration.Fields.Item _
'''        ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'''
'''    'Name or IP of Remote SMTP Server
'''    objMessage.Configuration.Fields.Item _
'''        ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com" 'SMTP SERVER of GMAIL must be inchanged
'''
'''    'Type of authentication, NONE, Basic (Base64 encoded), NTLM
'''    objMessage.Configuration.Fields.Item _
'''        ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
'''
'''    'Your UserID on the SMTP server
'''    objMessage.Configuration.Fields.Item _
'''        ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "swfranchise1@gmail.com" 'change this to yours
'''
'''    'Your password on the SMTP server
'''    objMessage.Configuration.Fields.Item _
'''        ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "franchise10" 'change this to yours
'''
'''    'Server port (typically 25 and 465 in SSL mode for Gmail)
'''    objMessage.Configuration.Fields.Item _
'''        ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465 'don't change this
'''
'''    'Use SSL for the connection (False or True)
'''    objMessage.Configuration.Fields.Item _
'''        ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
'''
'''    'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
'''    objMessage.Configuration.Fields.Item _
'''        ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
'''
'''    objMessage.Configuration.Fields.Update
'''
'''    '==End remote SMTP server configuration section==
'''
'''    objMessage.Send
'''
'''    If Err.Number <> 0 Then
'''            MsgBox Err.Description, 16, "Error"
'''            MsgBox "Could not send eMail !", 16, "Information"
'''        Else
''''        MsgBox "eMail sent", 64, "Information"
'''
'''    End If
'''
'''
'''    SendMail = True
'''
'''Leave:
'''    Exit Function
'''
'''ErrorHandler:
'''
'''    If Err = 32001 Then
'''        Resume Leave
'''    Else
'''        LogMsg frmEmail, Trim$(Error), ""
'''
'''    End If
'''
'''
'''End Function

Public Function GetEmailDetails(lID As Long, sInvoice As String, sDate As String, sClient As String, sTotalFee As String)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("Select * FROM tblDates INNER JOIN tblCLients ON tblDates.ClientID = tblClients.ID WHERE tblDates.ID = " & Trim$(lID))
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        sInvoice = gbRegion & "_" & Trim$(rs("InvNumber"))
        sDate = Trim$(Format(rs("on"), "dd/mmm/yyyy"))
        sClient = rs("txtName") & ""
        sTotalFee = Trim$(rs("INVTotal") & "")
        
        GetEmailDetails = True
        
    
    End If
    
    rs.Close

CleanExit:
    
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
'    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetEmailDetails") Then Resume 0
    Resume 0

End Function

Public Function fieldsok()

    If labelClientName.Caption <> "" Then
        If labelClientAddress.Caption <> "" Then
            If IsDate(labelDate.Caption) Then
                If labelFranRegion.Caption <> "" Then
                    If IsDate(lblDate) Then
                        If txtDescription.Text <> "" Then
                            If MoneyOk() Then
                                fieldsok = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

End Function

Public Function MoneyOk()

    If Val(txtAmount) = 0 Then
        If MsgBox("Fee is 0. Is this correct?", vbDefaultButton2 + vbYesNo + vbQuestion, "0 Fee") = vbNo Then
            MsgBox "A Regular Fee can be setup for this Client in the Client detail screen"
            bSetFocus Me, "txtamount"
            Exit Function
        Else
            MoneyOk = True
        End If
    
    Else
        If Val(txtAmount) = Val(labelFee) Then
            If labelVat = Format(Val(labelFee) * Val(labelVatRate) / 100, "0.00") Then
                If labelTotal = Format(Val(labelFee) + Val(labelVat), "0.00") Then
                    MoneyOk = True
                Else
                    MsgBox "Check Final Calculation"
                End If
            Else
                MsgBox "Check Final Calculation"
            End If
        Else
            MsgBox "Fee not matching subtotal"
        End If
    End If

End Function

Public Function SetDefaultFee(lID As Long, sFee As String)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblClients")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lID
    If Not rs.NoMatch Then
    
        rs.Edit
        rs("tedFee") = Val(sFee)
        rs.Update
    
    End If
    
    rs.Close

CleanExit:
    
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
'    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("SetDefaultFee") Then Resume 0
    Resume 0

End Function

