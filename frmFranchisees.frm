VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmFranchisees 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11400
   Icon            =   "frmFranchisees.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmFranchisees.frx":1CCA
   ScaleHeight     =   6240
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MyCommandButton.MyButton btnClose 
      Height          =   255
      Left            =   10995
      TabIndex        =   0
      Top             =   45
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
   Begin VSFlex8LCtl.VSFlexGrid grdFran 
      Height          =   4425
      Left            =   210
      TabIndex        =   2
      Top             =   1500
      Width           =   10695
      _cx             =   18865
      _cy             =   7805
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16052193
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
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
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFranchisees.frx":838B
      ScrollTrack     =   0   'False
      ScrollBars      =   3
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
      ExplorerBar     =   0
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
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmFranchisees.frx":8439
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
      Left            =   270
      TabIndex        =   3
      Top             =   1185
      Width           =   10350
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stockwatch - Franchisees"
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
      TabIndex        =   1
      Top             =   45
      Width           =   2280
   End
End
Attribute VB_Name = "frmFranchisees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()

    Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me

End Sub

Private Sub Form_Load()


    gbOk = ShowFranchisees()
    
End Sub

Public Function ShowFranchisees()
Dim rs As Recordset

    Set rs = swDB.OpenRecordset("SELECT * FROM tblFranchisees ORDER BY NAME")
    If Not (rs.EOF And rs.BOF) Then
        
        rs.MoveFirst
        
        grdFran.Rows = 0
        
        Do
        
            grdFran.AddItem rs("Region") & vbTab & rs("Name") & vbTab & rs("Address") & vbTab & rs("Phone") & vbTab & rs("joined") & vbTab & rs("email")
            grdFran.RowData(grdFran.Rows - 1) = rs("ID") + 0
            rs.MoveNext
        
        Loop While Not rs.EOF
        
    End If
    
 '   grdFran.AutoSize 0, grdFran.Cols - 1
    
    
    rs.Close

Leave:
    Exit Function

ErrorHandler:
    
    MsgBox "Error: " & Trim$(Error)
    Resume Leave
    Resume 0

End Function

Private Sub grdFran_Click()

    If grdFran.Row <> -1 And grdFran.Row < grdFran.Rows Then
        frmFranDetails.lFranID = grdFran.RowData(grdFran.Row)
        frmFranDetails.Show vbModal
    End If

End Sub

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

