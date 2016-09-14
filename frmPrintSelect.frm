VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmPrintSelect 
   BorderStyle     =   0  'None
   Caption         =   "StockWatch Print Reports"
   ClientHeight    =   5100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4905
   Icon            =   "frmPrintSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPrintSelect.frx":1CCA
   ScaleHeight     =   5100
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkHeader 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00B4A0B8&
      Caption         =   "Include Report &Header"
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
      Left            =   885
      TabIndex        =   1
      Top             =   3675
      Width           =   2385
   End
   Begin VSFlex8LCtl.VSFlexGrid grdReps 
      Height          =   2295
      Left            =   270
      TabIndex        =   0
      Top             =   1170
      Width           =   4395
      _cx             =   7752
      _cy             =   4048
      Appearance      =   2
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12157534
      ForeColorSel    =   -2147483634
      BackColorBkg    =   8421504
      BackColorAlternate=   16051944
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   8
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPrintSelect.frx":6E02
      ScrollTrack     =   0   'False
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   1.2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   -1  'True
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
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
      BackColorFrozen =   12157534
      ForeColorFrozen =   12157534
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin MSComctlLib.ImageList imgList 
         Left            =   90
         Top             =   6600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   21
         ImageHeight     =   21
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrintSelect.frx":6EE4
               Key             =   "button"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrintSelect.frx":71DA
               Key             =   "inp"
               Object.Tag             =   "inp"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrintSelect.frx":74EB
               Key             =   "correct"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrintSelect.frx":768C
               Key             =   "yellowflag"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrintSelect.frx":77EE
               Key             =   "redflag"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrintSelect.frx":7968
               Key             =   "Tick"
            EndProperty
         EndProperty
      End
   End
   Begin MyCommandButton.MyButton btnClose 
      Height          =   255
      Left            =   4500
      TabIndex        =   2
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
   Begin MyCommandButton.MyButton cmdPrint 
      Height          =   495
      Left            =   3615
      TabIndex        =   4
      Top             =   3615
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
      Caption         =   "&Print"
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
   Begin MyCommandButton.MyButton cmdQuit 
      Height          =   495
      Left            =   3630
      TabIndex        =   5
      Top             =   4440
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
   Begin VB.Label lblReportDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Name                                           Select"
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
      Left            =   870
      TabIndex        =   3
      Top             =   840
      Width           =   4035
   End
End
Attribute VB_Name = "frmPrintSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()

    cmdQuit_Click

End Sub

Private Sub cmdPrint_Click()
                 
    gbOk = PrintReports(frmStockWatch.lblClient.Tag, Replace(frmStockWatch.labelFrom.Tag, "/", "-"), Replace(frmStockWatch.labelTo.Tag, "/", "-"), chkHeader.Value)
                 
    Unload Me

End Sub

Private Sub cmdQuit_Click()

    Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then Unload Me

End Sub

Private Sub Form_Load()

    SetupMenu

End Sub

Private Sub grdReps_Click()
    If grdReps.Col = 1 Then grdReps.Col = 0

End Sub
Public Function SetupMenu()
Dim iRow As Integer

    grdReps.RowData(0) = "A"
    grdReps.RowData(1) = "B"
    grdReps.RowData(2) = "C"
    grdReps.RowData(3) = "D"
    grdReps.RowData(4) = "E"
    grdReps.RowData(5) = "F"
    grdReps.RowData(6) = "G"
    grdReps.RowData(7) = "I"
    
    For iRow = 0 To 7
        grdReps.Cell(flexcpPicture, iRow, 0) = frmStockWatch.imgList.ListImages("button").Picture
    
        If GetSetting(App.Title, "Reports", App.Title & "Print " & grdReps.RowData(iRow)) = "-1" Then
            grdReps.Cell(flexcpChecked, iRow, 2) = True
        Else
            grdReps.Cell(flexcpChecked, iRow, 2) = False
        End If
    
        grdReps.Cell(flexcpPictureAlignment, iRow, 0) = flexPicAlignCenterCenter

    
    Next

    If GetSetting(App.Title, "Reports", App.Title & "Print H") = "-1" Then
        chkHeader.Value = 1

' Ver440
'    Else
'        chkHeader.Value = 0
    End If


End Function

Private Sub grdReps_KeyPress(KeyAscii As Integer)
        
    grdReps.Row = grdReps.FindRow(UCase(Chr$(KeyAscii)), , 0)
        
    setClearSelection grdReps.Row, 2
        
End Sub


Public Sub setClearSelection(iRow As Integer, iCol As Integer)
Dim sRep As String
    
    If iRow > -1 And iCol = 2 Then
        
        sRep = "Print " & Trim$(grdReps.RowData(iRow))
        
        If grdReps.Cell(flexcpChecked, iRow, iCol) = 2 Then
            grdReps.Cell(flexcpText, iRow, iCol) = True
            
            SaveSetting appname:=App.Title, Section:="Reports", Key:=App.Title & sRep, Setting:=-1
        
        Else
            grdReps.Cell(flexcpText, iRow, iCol) = False
            SaveSetting appname:=App.Title, Section:="Reports", Key:=App.Title & sRep, Setting:=0
        
        End If
    End If


End Sub

Private Sub grdReps_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    setClearSelection grdReps.Row, grdReps.Col

End Sub

Public Function PrintReports(sName As String, sFromDate As String, sToDate As String, bIncludeCover As Boolean)
Dim objfile As Object
Dim sReport As String
Dim iRep As Integer
Dim sDate As String

    On Error GoTo ErrorHandler
    
    Set objfile = CreateObject("Scripting.FileSystemObject")
    ' create object

    gbOk = TerminateWINWORD()
    
    Set WriteWord = New Word.Application
   
    sDate = sToDate
    ' ver 550 moved to here
    
    For iRep = 7 To 0 Step -1 ' Rows in Menu to pick up report names
   
        If grdReps.Cell(flexcpChecked, iRep, 2) = 1 Then
    
'##########
' here to fix date folder problem
' reverse order here to look for To date first before From date check

'---------------------------------------------------------------------------------------------
' Ver440 this is removed in this version... it should no longer be required to check for folders
' created with the 'from date' since all reports are now generated using the 'to date' as the folder name.
'
'            sReport = sDBLoc & "\" & Trim$(sName) & "\" & Trim$(sFromDate) & "\" & grdReps.Cell(flexcpTextDisplay, iRep, 1) & ".Doc"
'            sDate = sFromDate
'
'            If objfile.FileExists(sReport) Then
'            ' see if file already exists
'
'                WriteWord.Visible = False
'
'                WriteWord.Application.Documents.Open sReport, , vbTrue
'                ' open report
'
'                WriteWord.Application.NormalTemplate.Saved = True
'
'                WriteWord.Application.PrintOut -1
'
'            Else
'---------------------------------------------------------------------------------------------
                
                sReport = sDBLoc & "\" & Trim$(sName) & "\" & Trim$(sToDate) & "\" & grdReps.Cell(flexcpTextDisplay, iRep, 1) & ".Doc"
                
                If objfile.FileExists(sReport) Then
                ' see if file already exists

                    WriteWord.Visible = False
        
                    WriteWord.Application.Documents.Open sReport, , vbTrue
                    ' open report
            
                    WriteWord.Application.NormalTemplate.Saved = True
    
                    WriteWord.Application.PrintOut -1
            
                Else
            
                    MsgBox "Report " & sReport & " does not exist"
            
                    WriteWord.Application.Quit
'                    WriteWord.Application.Quit SaveChanges:=wdDoNotSaveChanges
                    Set WriteWord = Nothing
                    Set objfile = Nothing
                    GoTo CleanExit
                
                End If
'            End If
            
        End If
        
    Next

    If bIncludeCover Then
    

    ' ver 550 this added here so it wouldnt bomb with no date
        If IsDate(sDate) Then
            gbOk = PrintReportCover(sName, sDate)
        End If
    
    Else
' ver 2.1.0
' this added here cause its already done in PrintReportCover above
        
        WriteWord.Application.Quit

        Set WriteWord = Nothing
        Set objfile = Nothing
'==========
    End If

    Pause 3000
    
            

CleanExit:
    
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
 '   If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("PrintReports") Then Resume 0
    Resume CleanExit


End Function
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

