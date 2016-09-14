VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Merge - Select from List or Click Add Button"
   ClientHeight    =   10215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5250
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtItem 
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
      Left            =   120
      TabIndex        =   4
      Top             =   1050
      Width           =   2895
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4350
      TabIndex        =   3
      Top             =   9600
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   990
      Width           =   735
   End
   Begin VSFlex8LCtl.VSFlexGrid grdSelect 
      Height          =   8445
      Left            =   90
      TabIndex        =   0
      Top             =   1620
      Width           =   4065
      _cx             =   7170
      _cy             =   14896
      Appearance      =   1
      BorderStyle     =   1
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSelect.frx":0000
      ScrollTrack     =   -1  'True
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label labelItem 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   6
      Top             =   180
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Change Item:"
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
      Left            =   180
      TabIndex        =   5
      Top             =   810
      Width           =   2205
   End
   Begin VB.Label labelMsg 
      Alignment       =   2  'Center
      BackColor       =   &H002841B5&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "- Not Found"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   180
      Width           =   1155
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lPLUID As Long
Public bAdd As Boolean
Public bStop As Boolean
Public sFind As String
Public sItem As String
Public iPointToRow As Integer



Private Sub VSFlexGrid1_Click()


End Sub

Private Sub cmdAdd_Click()
    
    sItem = UCase$(txtItem.Text)
    
    bAdd = True
    Unload Me
    
End Sub

Private Sub cmdStop_Click()

    bStop = True

    Unload Me

End Sub

Private Sub Form_Activate()
    
    If grdSelect.Rows > 1 Then
    
        If grdSelect.FindRow(sFind, , 0, , False) > -1 Then
            grdSelect.TopRow = grdSelect.FindRow(sFind, , 0, , False)
        End If
    End If
    
    
    DoEvents

End Sub

Private Sub Form_Deactivate()
Stop
End Sub

Private Sub Form_Load()
    bAdd = False
    bStop = False

    lPLUID = 0

    gbOk = ShowPLUs(sFind)
    
    DoEvents
    
End Sub

Private Sub grdSelect_Click()

    lPLUID = grdSelect.RowData(grdSelect.Row)
    Unload Me
    
End Sub
Private Function ShowPLUs(sFd As String)
Dim rs As Recordset
Dim sSql As String

        
    sSql = "WHERE chkActive = true"
    
    Screen.MousePointer = 11
    On Error GoTo ErrorHandler
    
    grdSelect.Rows = 0
    
    Set rs = SWdb.OpenRecordset("Select * FROM tblPLUs " & sSql & " ORDER BY txtDescription", dbOpenSnapshot)
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do
            grdSelect.AddItem rs("txtDescription") & vbTab & rs("Other")
            grdSelect.RowData(grdSelect.Rows - 1) = rs("ID") + 0
            
            rs.MoveNext
        Loop While Not rs.EOF
    End If
    
    DoEvents
    
    ShowPLUs = True

CleanExit:
    DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    Screen.MousePointer = 0
    
    Exit Function

ErrorHandler:
 '   If CheckDBError("ShowPLUs") Then Resume 0
    Resume CleanExit


End Function

Private Sub txtItem_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))


End Sub
