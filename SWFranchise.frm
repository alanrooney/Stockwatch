VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFranchise 
   BorderStyle     =   0  'None
   ClientHeight    =   9780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9750
   Icon            =   "SWFranchise.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "SWFranchise.frx":1CCA
   ScaleHeight     =   9780
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView Cal 
      Height          =   2820
      Left            =   3975
      TabIndex        =   17
      Top             =   1965
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
      StartOfWeek     =   27787266
      TitleBackColor  =   13342061
      TitleForeColor  =   16777215
      CurrentDate     =   39972
   End
   Begin VB.FileListBox FilTransfer 
      Height          =   675
      Left            =   195
      Pattern         =   "*.csv"
      TabIndex        =   4
      Top             =   8925
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VSFlex8LCtl.VSFlexGrid grdFran 
      Height          =   6345
      Left            =   360
      TabIndex        =   22
      Top             =   2760
      Width           =   9045
      _cx             =   15954
      _cy             =   11192
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16052193
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16052193
      GridColor       =   -2147483633
      GridColorFixed  =   16052193
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   60
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"SWFranchise.frx":93C3
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSMask.MaskEdBox tedFrom 
      Height          =   375
      Left            =   3975
      TabIndex        =   18
      Top             =   1575
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   16777215
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
   Begin MyCommandButton.MyButton btnCallFrom 
      Height          =   360
      Left            =   4965
      TabIndex        =   10
      Top             =   1590
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "SWFranchise.frx":9481
      BackColorDown   =   15133676
      TransparentColor=   14215660
      Caption         =   ""
      DepthEvent      =   1
      PictureDisabled =   "SWFranchise.frx":C886
      ShowFocus       =   -1  'True
   End
   Begin VB.CheckBox chkAllDates 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F4EFE1&
      Caption         =   "All"
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
      Left            =   4710
      TabIndex        =   7
      Top             =   1200
      Width           =   540
   End
   Begin VB.CheckBox chkAllRegions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F4EFE1&
      Caption         =   "All"
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
      Left            =   2010
      TabIndex        =   5
      Top             =   1200
      Width           =   540
   End
   Begin VB.ComboBox cboRegions 
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
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1590
      Width           =   1605
   End
   Begin MyCommandButton.MyButton btnCalTo 
      Height          =   360
      Left            =   6855
      TabIndex        =   11
      Top             =   1605
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "SWFranchise.frx":CBA8
      BackColorDown   =   15133676
      TransparentColor=   14215660
      Caption         =   ""
      DepthEvent      =   1
      PictureDisabled =   "SWFranchise.frx":FFAD
      ShowFocus       =   -1  'True
   End
   Begin MyCommandButton.MyButton btnClose 
      Height          =   255
      Left            =   9360
      TabIndex        =   12
      Top             =   75
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
   Begin MyCommandButton.MyButton btnCheckForFiles 
      Height          =   300
      Left            =   8880
      TabIndex        =   13
      Top             =   570
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   529
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
      Caption         =   ">>"
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
   Begin MyCommandButton.MyButton btnGo 
      Height          =   360
      Left            =   8640
      TabIndex        =   15
      Top             =   1605
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   635
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
      Caption         =   "Go >"
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
   Begin MSMask.MaskEdBox tedTo 
      Height          =   375
      Left            =   5880
      TabIndex        =   19
      Top             =   1590
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   661
      _Version        =   393216
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
   Begin MyCommandButton.MyButton btnManage 
      Height          =   360
      Left            =   435
      TabIndex        =   20
      Top             =   570
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   635
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
      Caption         =   "Franchisees"
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
   Begin VB.Label labelINVReceived 
      BackStyle       =   0  'Transparent
      Caption         =   "New Invoice(s) received"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   2640
      TabIndex        =   21
      Top             =   630
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stockwatch - Franchise Management Program"
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
      Left            =   495
      TabIndex        =   16
      Top             =   75
      Width           =   4110
   End
   Begin VB.Label lblCheckForFiles 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check For New Invoice Records"
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
      Left            =   5955
      TabIndex        =   14
      Top             =   600
      Width           =   2880
   End
   Begin VB.Label lblTo 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5580
      TabIndex        =   9
      Top             =   1635
      Width           =   1485
   End
   Begin VB.Label lblFrom 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3465
      TabIndex        =   8
      Top             =   1635
      Width           =   1485
   End
   Begin VB.Label lblDates 
      BackStyle       =   0  'Transparent
      Caption         =   "Dates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3960
      TabIndex        =   6
      Top             =   1200
      Width           =   1485
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6525
      TabIndex        =   3
      Top             =   9315
      Width           =   975
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
      Height          =   240
      Left            =   7515
      TabIndex        =   2
      Top             =   9300
      Width           =   1140
   End
   Begin VB.Label lblRegion 
      BackStyle       =   0  'Transparent
      Caption         =   "Regions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1050
      TabIndex        =   0
      Top             =   1185
      Width           =   1485
   End
End
Attribute VB_Name = "frmFranchise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bFromDate As Boolean

Private Sub btnCallFrom_Click()
    If Cal.Visible = True Then
        Cal.Visible = False
    Else
        If IsDate(tedFrom) Then
            Cal.Value = tedFrom
        Else
            Cal.Value = Format(Now, "dd/mm/yy")
        End If
        bFromDate = True
        Cal.Top = tedFrom.Top + tedFrom.Height
        Cal.Left = tedFrom.Left
        Cal.Visible = True
    End If

End Sub

Private Sub btnCalTo_Click()
    If Cal.Visible = True Then
        Cal.Visible = False
    Else
        If IsDate(tedTo) Then
            Cal.Value = tedTo
        Else
            Cal.Value = Format(Now, "dd/mm/yy")
        End If
        bFromDate = False
        Cal.Top = tedTo.Top + tedTo.Height
        Cal.Left = tedTo.Left
        Cal.Visible = True
    End If
    

End Sub

Private Sub btnCheckForFiles_Click()

    labelINVReceived.Visible = False
    ' clear msg first
    
    gbOk = CheckForNewFranchisee()
    ' To come from my setup program
    ' Includes Name, Address, Email and Expiry Date
    
    gbOk = GetRegions(Me)
    ' get regions again incase a new one has arrived
    
    gbOk = CheckForInvoiceFiles()

End Sub

Private Sub btnClose_Click()

    End
    

End Sub

Private Sub btnGo_Click()

    If DatesOk() Then
    
        gbOk = DoFranchiseReport()

    End If

End Sub

Private Sub btnManage_Click()

    frmFranchisees.Show vbModal

End Sub

'Private Sub btnPaid_Click()
'
'    If btnPaid.ToggleValue Then
'        btnUnPaid.ToggleValue = False
'    End If
'
'
'End Sub
'
'Private Sub btnUnPaid_Click()
'    If btnUnPaid.ToggleValue Then
'        btnPaid.ToggleValue = False
'    End If
'
'End Sub

Private Sub Cal_DateClick(ByVal DateClicked As Date)
    
    If bFromDate Then
         tedFrom.Text = Format(Cal.Value, "dd/mm/yy")
    Else
         tedTo.Text = Format(Cal.Value, "dd/mm/yy")
    End If
    
    Cal.Visible = False

End Sub

Private Sub cboRegions_Click()

    chkAllRegions.Value = 0
    
End Sub

Private Sub cboRegions_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub chkAllDates_Click()

    If chkAllDates.Value = 1 Then
        tedFrom.ForeColor = &HC0C0C0
        tedTo.ForeColor = &HC0C0C0
    Else
        tedFrom.ForeColor = &H80000008
        tedTo.ForeColor = &H80000008
    End If

End Sub

Private Sub chkAllRegions_Click()

    If chkAllRegions.Value = 1 Then
        cboRegions.ForeColor = &HC0C0C0
    Else
        cboRegions.ForeColor = &H80000008
    End If


End Sub

'Private Sub chkAllStatus_Click()
'
'
'    If chkAllStatus.Value = 0 Then
'
'        btnPaid.Enabled = True
'        btnUnPaid.Enabled = True
'    Else
'        btnPaid.Enabled = False
'        btnUnPaid.Enabled = False
'
'    End If
'
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    ' return pressed force focus to next available object in tabbing order
        gbOk = GotoNextControl(Me, 0)
    
'    ElseIf KeyAscii = 27 Then
'        btnClose_Click
'
    End If


End Sub

Private Sub Form_Load()

    gbOk = GetXFERLocation(gbsw1)
    
    If CurDir$ = "C:\Program Files\Microsoft Visual Studio\VB98" Then
    ' development environment
    
        gbDropBox = Replace(gbsw1, "\SW2", "")
    
    Else
    ' live
    
        gbDropBox = Replace(gbsw1, "\SW1", "")
    
    End If
    
    If gbOpenDB(Me) Then
    ' now open db
    
        gbOk = GetRegions(Me)
    
        sHORegion = Right(gbsw1, 3)
        
        btnCheckForFiles_Click

    End If
    
End Sub

Public Function CheckForInvoiceFiles()
Dim rs As Recordset
Dim iCnt As Integer
Dim sRecvFile As String
Dim sDropBoxLoc As String
Dim iRegion As Integer

    On Error GoTo ErrorHandler
    
    For iRegion = 0 To cboRegions.ListCount - 1
    
        On Error GoTo ProblemAccessingRegionDropBoxFolder
        
        sDropBoxLoc = Replace(gbsw1, sHORegion, cboRegions.List(iRegion))
        ' Get each regions dropbox location
        
        FilTransfer.Path = sDropBoxLoc
        FilTransfer.Refresh
        ' point to each region dropbox inturn
        
        For iCnt = 0 To FilTransfer.ListCount - 1
        ' loop throught files to be received/processed
        
            If Left(FilTransfer.List(iCnt), 4) <> "Ack_" And Left(FilTransfer.List(iCnt), 4) <> "NEW_" Then
            ' Ignore Acknowledge and New files
            
                sRecvFile = FilTransfer.List(iCnt)
            
                If CopyFileFromDropBox(sDropBoxLoc & "\" & sRecvFile, sDBLoc & sRecvFile) Then
                    
                    If UpdateDBandAck(sDBLoc, sRecvFile, sDropBoxLoc) Then
                        
                        labelINVReceived.Visible = True
    
                    End If
                End If
            
            End If
            
        Next
    Next
    
    
Leave:
    Exit Function

ErrorHandler:
    
    If Err = 76 Then Resume Leave
    ' path not found
    
    MsgBox "Error: " & Trim$(Error) & " - Problem updating the Franchise Database"
    Resume Leave
    Resume 0
    
ProblemAccessingRegionDropBoxFolder:
    MsgBox "Error: " & Trim$(Error) & " - Problem updating the Franchise Database"
    Resume Leave
    

End Function
Public Function CopyFileFromDropBox(sSource As String, sDest As String)
Dim objFile As Object

    On Error GoTo RenameError
            
    Set objFile = CreateObject("Scripting.FileSystemObject")
    
    If objFile.FileExists(sDest) Then
    ' See if destination file already exists and remove it
    
        Kill sDest
    End If
    
    Name sSource As sDest

'    Name sSource As "test.csv"

    CopyFileFromDropBox = True

Leave:
    Exit Function

RenameError:
    MsgBox "Error: " & Trim$(Error) & " - Problem copying file from " & sSource & " to " & sDest
    Resume Leave

End Function

Public Function UpdateDBandAck(sLocation As String, sINVFile As String, sDropBox As String)

Dim lRCV As Long
Dim lAck As Long
Dim sData As String
Dim sAckData As String
Dim rs As Recordset
Dim sTemp As String
Dim lInvoiceNo As Long
Dim dtDate As Date
Dim sClient As String
Dim curFee As Currency
Dim sInvoice As String
Dim sRegion As String
Dim lID As Long

    On Error GoTo ErrorHandler
    
    lRCV = FreeFile
    
    
    On Error GoTo CantReadInvoiceFile
   
    
    Open sLocation & sINVFile For Input As #lRCV   ' Open file for input.
    ' open the booking file passed
    
    Line Input #lRCV, sData
    Close #lRCV
    ' Read all fields - there's only one record
    
    On Error GoTo ProblemDecryptingInvoiceFile
    
    sAckData = Decrypt(sData, sKey)
    ' unbundle it

    sData = Replace(sAckData, "@@", vbCrLf)
    ' fix up address lines
    
    sRegion = Mid(sData, InStr(1, sData, "<Region>") + 8, InStr(1, sData, "/<Region>") - InStr(1, sData, "<Region>") - 8)
    sInvoice = Mid(sData, InStr(1, sData, "<InvNumber>") + 11, InStr(1, sData, "/<InvNumber>") - InStr(1, sData, "<InvNumber>") - 11)
    
    lID = 0 ' Default to New Invoice
    
    ' FIRST MAKE SURE INVOICE RECEIVED IS NOT ALREADY ON THE SYSTEM
    
    
    On Error GoTo ProblemCheckingIfInvoiceSummaryFileOnSystem
    
    Set rs = swDB.OpenRecordset("SELECT * FROM tblStockTakes WHERE Region = '" & sRegion & "' AND InvoiceNumber = " & sInvoice)
    If Not (rs.EOF And rs.BOF) Then
    
        ' INVOICE ALREADY RECEIVED!
        
        If MsgBox("Invoice already received: " & rs("Region") & " " & rs("InvoiceNumber") & " " & rs("Date") & " " & rs("ClientName") & " " & rs("TotalFee") & ". Overwrite with new one received?", vbDefaultButton2 + vbYesNo + vbQuestion, "Invoice Already Received") = vbYes Then
            lID = rs("ID") + 0
        End If
    
    End If
    rs.Close
    
    On Error GoTo ProblemUpdatingSummaryFileIntoDB
    
    Set rs = swDB.OpenRecordset("tblStockTakes")
    rs.Index = "PrimaryKey"
    If lID <> 0 Then
        rs.Seek "=", lID
        rs.Edit
    Else
        rs.AddNew
    End If
    
    rs("Region") = Mid(sData, InStr(1, sData, "<Region>") + 8, InStr(1, sData, "/<Region>") - InStr(1, sData, "<Region>") - 8)
    rs("InvoiceNumber") = Mid(sData, InStr(1, sData, "<InvNumber>") + 11, InStr(1, sData, "/<InvNumber>") - InStr(1, sData, "<InvNumber>") - 11)
    rs("Date") = Mid(sData, InStr(1, sData, "<Date>") + 6, InStr(1, sData, "/<Date>") - InStr(1, sData, "<Date>") - 6)
    rs("ClientName") = Mid(sData, InStr(1, sData, "<Name>") + 6, InStr(1, sData, "/<Name>") - InStr(1, sData, "<Name>") - 6)
    rs("TotalFee") = Mid(sData, InStr(1, sData, "<Total>") + 7, InStr(1, sData, "/<Total>") - InStr(1, sData, "<Total>") - 7)
    
    rs.Update
    rs.Close
            
    
    ' CREATE ACK FILE HERE
    
    On Error GoTo ProblemCreatingACKFile
    
    lAck = FreeFile
    
    sData = Encrypt(sAckData & "\<Ack>", sKey)
    
    Open sLocation & "Ack_" & sINVFile For Output As #lAck ' Open file for input.
    ' open the booking file passed
    
    Print #lAck, sData
    Close #lAck
    
    ' COPY OUT ACK FILE
    
    On Error GoTo ProblemCopyingACKFileToDropBoxClientLocation
    
    gbOk = CopyFileToDropBox(sLocation & "Ack_" & sINVFile, sDropBox & "\Ack_" & sINVFile)
    
    
    UpdateDBandAck = True

CleanExit:
    DBEngine.Idle dbRefreshCache
     ' Release unneeded DB locks
    
    Exit Function

ErrorHandler:
    If CheckDBError("UpdateDBandAck") Then Resume 0
    Resume CleanExit
    Resume 0

CantReadInvoiceFile:
    MsgBox Trim$(Error) & "- Cant Read Invoice File"
    Resume CleanExit

ProblemDecryptingInvoiceFile:
    MsgBox Trim$(Error) & "- Problem Decrypting Invoice File"
    Resume CleanExit

ProblemCheckingIfInvoiceSummaryFileOnSystem:
    MsgBox Trim$(Error) & "- Problem Checking If Invoice Summary File On System"
    Resume CleanExit

ProblemUpdatingSummaryFileIntoDB:
    MsgBox Trim$(Error) & "- Problem Updating Summary File Into DB"
    Resume CleanExit

ProblemCreatingACKFile:
    MsgBox Trim$(Error) & "- Problem Creating ACK File"
    Resume CleanExit

ProblemCopyingACKFileToDropBoxClientLocation:
    MsgBox Trim$(Error) & "- Problem Copying ACK File To DropBox Client Location"
    Resume CleanExit

End Function

Function gbOpenDB(mainfrm As Form) As Boolean
    
    On Error GoTo ErrorHandler
    
    sDBLoc = "" & GetSetting(App.Title, "DB", App.Title & "DB") & ""
    ' get the DB Location from the registry
    
    ' remove the extension since customer windows expl might not be showing extensions
    
    If sDBLoc = "" Then
        sDBLoc = InputBox("Enter Database Location (C:\" & App.Title & ")", "Invalid Database Location: " & sDBLoc)
        If sDBLoc = "exit" Then End
        SaveSetting appname:=App.Title, Section:="DB", Key:=App.Title & "DB", Setting:=sDBLoc
    End If
    
OpenDB:
    Set swDB = OpenDatabase("" & sDBLoc & "\" & App.Title & ".mdb", False, False)
'    LogMsg frmSubMan, "DataBase Opened", " File: " & sDBLoc
    gbOpenDB = True

CleanExit:
    Exit Function
    
ErrorHandler:
    
    sDBLoc = InputBox("Enter Database Location (C:\" & App.Title & "\" & App.Title & ".mdb)", "Invalid Database Location: " & sDBLoc)
    
    
    If sDBLoc = "exit" Then End
    SaveSetting appname:=App.Title, Section:="DB", Key:=App.Title & "DB", Setting:=sDBLoc
    Resume OpenDB
    
End Function

Public Function CheckDBError(sSection As String)
  
  Dim endofpause As Double
  Dim errorloop As Error
  
  CheckDBError = True
  ' default to true
  ' only false when retries are exhausted
      
  If (Err > 2999 And Err < 4000) Or Err = 75 Or Err = 55 Or Err = 57 Or Err = 71 Or Err = 76 Then
    ' if its a database error...
    iErrCount = iErrCount + 1
      
'    LogMsg mainfrm, "Database Locked by another user, Waiting...", ""
    If iErrCount = 20 Then
'      LogMsg mainfrm, "", "Database locked by another user " & Error
      CheckDBError = False
            
      Exit Function
      ' wait exhausted so leave routine with the bad news...
    End If
  
    endofpause# = Timer + 1
    
    Do
     
    
    
    Loop While Timer < endofpause#
    ' wait one
  
  Else
    SaveDebug sSection, "Error #" + Str$(Err) + " [" + Trim$(Error) + "] occured"
    
    MsgBox sSection & " Error #" & Str$(Err) & " [" & Trim$(Error) & "] occured"
    End
  End If

End Function
Public Function GotoNextControl(frmname As Form, iStartCtl As Integer)
Dim iNextTabIndex As Integer
        
  On Error GoTo ErrorHandler
    
    If iStartCtl > 0 Then
        iNextTabIndex = iStartCtl
        
    ElseIf frmname.ActiveControl.TabIndex = Screen.ActiveForm.Count - 1 Then
        iNextTabIndex = iStartCtl
    Else
        iNextTabIndex = frmname.ActiveControl.TabIndex + 1
    End If
    ' if we're at the last control then point to the first control
    ' otherwise point to the next control
        
        
    On Error GoTo skipThisCtrl
GetNextControl:
    For gbCnt = 0 To Screen.ActiveForm.Count - 1
        
      If TypeName(Screen.ActiveForm.Controls(gbCnt)) <> "Skinner" And TypeName(Screen.ActiveForm.Controls(gbCnt)) <> "Line" And TypeName(Screen.ActiveForm.Controls(gbCnt)) <> "Image" And TypeName(Screen.ActiveForm.Controls(gbCnt)) <> "CommonDialog" Then
         If frmname.Controls(gbCnt).TabIndex = iNextTabIndex Then
        
            If Screen.ActiveForm.Controls(gbCnt).Enabled And Screen.ActiveForm.Controls(gbCnt).Visible Then
            
                Select Case TypeName(Screen.ActiveForm.Controls(gbCnt))
                ' we only want to set focus to controls that will allow it
                
                    Case "ListBox", "TextBox", "ComboBox", "CommandButton", "MaskEdBox", "OptionButton", "RichTextBox", "VSFlexGrid", "DriveListBox", "CheckBox"
                     If Screen.ActiveForm.Controls(gbCnt).TabStop Then
                     ' make sure tab stop is enabled for the control...
                     
                        Screen.ActiveForm.Controls(gbCnt).SetFocus
                     Else
                         iNextTabIndex = iNextTabIndex + 1
                         GoTo GetNextControl
                     End If
                     Exit For
                    
                    Case Else
                     iNextTabIndex = iNextTabIndex + 1
                     GoTo GetNextControl
                     ' yugh! but it works
                     
                End Select
                    
            Else
                iNextTabIndex = iNextTabIndex + 1
                GoTo GetNextControl
            End If
        
        End If
      End If
      
NextCtrl:
    Next gbCnt
  
CleanExit:
    Exit Function
        
ErrorHandler:
    Resume CleanExit
Resume 0

skipThisCtrl:
    Resume NextCtrl
    
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


Private Sub grdFran_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    If Col <> grdFran.ColIndex("paid") Then Cancel = True

End Sub

Private Sub grdFran_Click()

    If grdFran.Col = grdFran.ColIndex("paid") Then
    
        gbOk = UpdateFranchise(grdFran.RowData(grdFran.Row), grdFran.Cell(flexcpChecked, grdFran.Row, grdFran.Col))
    
    End If

End Sub

Private Sub MyButton1_Click()

End Sub

Private Sub tedFrom_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub tedTo_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Public Function DoFranchiseReport()
Dim rs As Recordset
Dim sSql As String
Dim sWhereAnd As String
Dim curTotal As Currency

    On Error GoTo ErrorHandler
    
    sWhereAnd = " WHERE "
    grdFran.Rows = 1
    
    'REGIONS
    If chkAllRegions.Value = 1 Or cboRegions = "" Then
    
        sSql = ""
        
    ElseIf cboRegions.ListIndex <> -1 Then
        sSql = sWhereAnd & "region = '" & Trim$(cboRegions) & "'"
        sWhereAnd = " AND "
    End If
    
    'DATES
    If chkAllDates.Value = 1 Then
    
    ElseIf IsDate(tedFrom) And IsDate(tedTo) Then
        sSql = sSql & sWhereAnd & "[Date] >= #" & Format(tedFrom, "mm/dd/yy") & "# AND [Date] <= #" & Format(tedTo, "mm/dd/yy") & "#"
        sWhereAnd = " AND "
    End If
    
    curTotal = 0
    
    Set rs = swDB.OpenRecordset("SELECT * from tblStockTakes " & sSql & " Order By [Date] ASC")
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do
        
            grdFran.AddItem rs("Region") & vbTab & rs("InvoiceNumber") & vbTab & rs("Date") & vbTab & rs("ClientName") & vbTab & Format(rs("TotalFee"), "0.00") & vbTab & rs("Paid")
            grdFran.RowData(grdFran.Rows - 1) = rs("ID") + 0
            
            curTotal = curTotal + rs("TotalFee")
            
            rs.MoveNext
        Loop While Not rs.EOF
    End If
    
    rs.Close
    
    labelTotal.Caption = Format(curTotal, "Currency")

Leave:
    Exit Function

ErrorHandler:
    
    MsgBox "Error: " & Trim$(Error)
    Resume Leave
    Resume 0
    
End Function

Public Function DatesOk()

    If chkAllDates.Value = 1 Then
        DatesOk = 1
    ElseIf IsDate(tedFrom) Then
        
        If IsDate(tedTo) Then
        
            If DateValue(tedFrom) <= DateValue(tedTo) Then
                DatesOk = 1
            Else
            
                MsgBox "From Date must be Older than To Date"
                bSetFocus Me, "tedFrom"
            End If
        Else
            bSetFocus Me, "tedTo"
        End If
    Else
        bSetFocus Me, "tedFrom"
    End If

End Function

Public Function UpdateFranchise(lID As Long, iPaid As Integer)
Dim rs As Recordset

    On Error GoTo ErrorHandler

    Set rs = swDB.OpenRecordset("tblFranchise")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lID
    If Not rs.NoMatch Then
        rs.Edit
        If iPaid = 2 Then
            rs("paid") = True
        Else
            rs("paid") = False
        End If
        
        rs.Update
    End If
    
    rs.Close

Leave:
    Exit Function

ErrorHandler:
    
    MsgBox "Error: " & Trim$(Error)
    Resume Leave
    Resume 0
    
End Function
'''Function GetXferLocationFromMainStockwatchDB(sXferLocation) As Boolean
'''Dim rs As Recordset
'''
'''    On Error GoTo ErrorHandler
'''
'''    sDBLoc = "" & GetSetting("StockWatch", "DB", "StockWatch" & "DB") & ""
'''    ' get the DB Location from the registry
'''
'''OpenDB:
'''    Set swDB = OpenDatabase("" & sDBLoc & "\Stockwatch.mdb", False, False, ";PWD=fran2012")
'''
'''
'''    Set rs = swDB.OpenRecordset("tblFranchisee")
'''    If Not (rs.EOF And rs.BOF) Then
'''        rs.MoveFirst
'''        sXferLocation = rs("XferLocation") & ""
'''    End If
'''    rs.Close
'''
'''    swDB.Close
'''
'''    GetXferLocationFromMainStockwatchDB = True
'''
'''CleanExit:
'''    Exit Function
'''
'''ErrorHandler:
'''
'''        MsgBox Error
'''
'''    Resume CleanExit
'''
'''End Function

Public Function CheckForNewFranchisee()
Dim rs As Recordset
Dim objFile As Object
Dim filenum As Long
Dim sLic As String
Dim sTestExpiry As String
Dim sRegion As String

    Set objFile = CreateObject("Scripting.FileSystemObject")
    
    If objFile.FileExists(gbsw1 & "\NEW_Maint.csv") Then
    ' See if theres a new franchise file present

        On Error Resume Next
        Kill sDBLoc & "\NEW_Maint.csv"
        
        On Error GoTo ProblemCopyingFile
        Name gbsw1 & "\NEW_Maint.csv" As sDBLoc & "\NEW_Maint.csv"
        ' Copy file in from dropbox and remove old file at same time
        
        filenum = FreeFile
        Open sDBLoc & "\NEW_Maint.csv" For Input As #filenum
        ' open it
        
        Input #filenum, sLic
        Close #filenum
        
        sLic = Decrypt(sLic, sKey)
        ' unbundle it

        sLic = Replace(sLic, "@@", vbCrLf)
        ' fix up address lines

        sRegion = Mid(sLic, InStr(1, sLic, "<Region>") + 8, InStr(1, sLic, "/<Region>") - InStr(1, sLic, "<Region>") - 8)
        
        Set rs = swDB.OpenRecordset("SELECT ID FROM tblFranchisees WHERE Region = '" & sRegion & "'")
        If (rs.EOF And rs.BOF) Then
        
            Set rs = swDB.OpenRecordset("tblFranchisees")
            rs.Index = "PrimaryKey"
            
            rs.AddNew
        
            rs("Region") = Mid(sLic, InStr(1, sLic, "<Region>") + 8, InStr(1, sLic, "/<Region>") - InStr(1, sLic, "<Region>") - 8)
            rs("Name") = Mid(sLic, InStr(1, sLic, "<Name>") + 6, InStr(1, sLic, "/<Name>") - InStr(1, sLic, "<Name>") - 6)
            rs("Address") = Mid(sLic, InStr(1, sLic, "<Address>") + 9, InStr(1, sLic, "/<Address>") - InStr(1, sLic, "<Address>") - 9)
            rs("Phone") = Mid(sLic, InStr(1, sLic, "<Phone>") + 7, InStr(1, sLic, "/<Phone>") - InStr(1, sLic, "<Phone>") - 7)
            rs("Email") = Mid(sLic, InStr(1, sLic, "<Email>") + 7, InStr(1, sLic, "/<Email>") - InStr(1, sLic, "<Email>") - 7)
            rs("Joined") = Mid(sLic, InStr(1, sLic, "<Joined>") + 8, InStr(1, sLic, "/<Joined>") - InStr(1, sLic, "<Joined>") - 8)
            rs("Expiry") = Mid(sLic, InStr(1, sLic, "<Expiry>") + 8, InStr(1, sLic, "/<Expiry>") - InStr(1, sLic, "<Expiry>") - 8)
            rs("Days") = Mid(sLic, InStr(1, sLic, "<Days>") + 6, InStr(1, sLic, "/<Days>") - InStr(1, sLic, "<Days>") - 6)
            rs("Warn") = Mid(sLic, InStr(1, sLic, "<Warn>") + 6, InStr(1, sLic, "/<Warn>") - InStr(1, sLic, "<Warn>") - 6)
            
            rs.Update
        
            MsgBox "New Franchise Details Received", vbOKOnly + vbInformation, "Stockwatch"
        
            Kill sDBLoc & "\NEW_Maint.csv"
        
        Else
            MsgBox "New Franchise Details Received But not added, Region " & sRegion & " already present"
    
        End If
    
    
    
    End If

Leave:
    Exit Function
    
ProblemCopyingFile:
    MsgBox "Problem Copying New Franchise File From Dropbox"
    Resume Leave
    
    
End Function

Public Function GetXFERLocation(sXFerLoc As String)
    
    sXFerLoc = GetSetting("StockWatch", "XFER", "Location")
    
    If sXFerLoc = "" Then
    ' either blank
    
        MsgBox "No Xfer Location specified. Run StockWatch again to reset it"
        
    End If
    
End Function

'''Public Sub CreateAckFile(sACKFile As String)
'''Dim lAck As Long
'''Dim sData As String
'''Dim filenum As Long
'''
'''    lAck = FreeFile
'''
'''    Open sACKFile For Output As lAck   ' Open file for input.
'''    ' open
'''
'''    Line Input #lAck, sData
'''    Close #lAck
'''    ' Read all fields - there's only one record
'''
'''    sData = Decrypt(sData, sKey)
'''    ' unbundle it
'''
'''    sData = sData & "_Ack"
'''
'''    ' ENCRYPT again
'''    sEncrypt = Encrypt(sData, sKey)
'''
'''    ' CREATE THE FILE
'''    filenum = FreeFile
'''    Open sDBLoc & "\" & gbRegion & "_" & rs("InvNumber") & ".csv" For Output As #filenum
'''
'''    Print #filenum, sEncrypt
'''    Close #filenum
'''
'''
'''End Sub
