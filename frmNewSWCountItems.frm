VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmNewSWCountItems 
   BackColor       =   &H00F2E2C1&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7485
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MyCommandButton.MyButton cmdOk 
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Top             =   5805
      Width           =   1380
      _ExtentX        =   2434
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
      TransparentColor=   15917761
      Caption         =   "&Proceed"
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
   Begin MyCommandButton.MyButton cmdQuit 
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   5820
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
      TransparentColor=   15917761
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
   Begin VSFlex8LCtl.VSFlexGrid grdNew 
      Height          =   5050
      Left            =   120
      TabIndex        =   0
      Top             =   510
      Width           =   7290
      _cx             =   12859
      _cy             =   8908
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
      BackColorFixed  =   16571070
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16777215
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   360
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmNewSWCountItems.frx":0000
      ScrollTrack     =   0   'False
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
   Begin MyCommandButton.MyButton btnClose 
      Height          =   255
      Left            =   7110
      TabIndex        =   3
      Top             =   105
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
      TransparentColor=   15917761
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Stockwatch - New Items from SWCount (Tablet)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   135
      TabIndex        =   4
      Top             =   105
      Width           =   6030
   End
End
Attribute VB_Name = "frmNewSWCountItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sNewFileName As String
Public sDrive As String
Public bDeleteNewItemsFile As Boolean



Private Sub btnClose_Click()
    
    cmdQuit_Click

End Sub

Private Sub cmdOk_Click()


    'KILL NEXT FILE
    Dim sFile, sf
    
    Dim objfile
    
    Set objfile = CreateObject("Scripting.FileSystemObject")
    ' create object
    
    ' DELETE THE ITEMS FILE
    
    
    On Error Resume Next
        
    If bDeleteNewItemsFile Then
        
        If objfile.FileExists(sNewFileName) Then
            
            Set sFile = objfile.GetFile(sNewFileName)
            sf = sFile.Delete
            ' kill old
        End If
        
    Else


    ' delete new items file
    
    ' call procedure in main
    
        If ImportTabletFile(lSelClientID, sDrive) Then
                            
           If DeleteImportedFiles(lSelClientID, sDrive) Then
           
                MsgBox "Import Complete"
            End If
        End If
    
    End If
    

    cmdQuit_Click




    








End Sub

Private Sub cmdQuit_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    
    Me.Left = 0
    Me.Top = 700
    
    AlwaysOnTop frmNewSWCountItems, True


    gbOk = getNewItems()

'    setformsize



End Sub

'The following example shows how to make a form stay on top of other forms.
'To try this example, paste the SetWindowPos Function declaration into the module level
'section of a BAS file. Place the remainder of the procedure in a BAS module also.
'To run the code, call the AlwaysOnTop sub from the form you want to stay on top as '
' shown in the Usage section of the comments within the procedure.


Sub AlwaysOnTop(FrmID As Form, OnTop As Integer)
    ' ===========================================
    ' Requires the following declaration
    ' For VB3:
    ' Declare Function SetWindowPos Lib "user" (ByVal h%, ByVal hb%, ByVal x%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer
    ' For VB4:
    ' ===========================================
    ' Usage:
    ' AlwaysOnTop Me, -1  ' To make always on top
    ' AlwaysOnTop Me, -2  ' To make NOT always on top
    ' ===========================================
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const flags = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    If OnTop = -1 Then
        OnTop = SetWindowPos(FrmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
    Else
        OnTop = SetWindowPos(FrmID.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
    End If
End Sub


Public Function getNewItems()
Dim lNewf As Long
Dim sIn As String
Dim sHdr As String
Dim sDates As String
Dim sItem As String
Dim sBar As String
Dim sCount As String
Dim sGroup As String
Dim sSize As String
Dim sFullWt As String
Dim sEmptyWt As String
Dim sIssueUnits As String
Dim sNotes As String
Dim sFull As String
Dim sOpen As String
Dim sWeight As String

Dim i As Integer

    On Error GoTo ErrorHandler
    
    lNewf = FreeFile
    Open sNewFileName For Input As #lNewf

    
    Line Input #lNewf, sIn
    If Left(sIn, 7) = "!Client" Then
        sHdr = Mid(sIn, InStr(9, sIn, ",") + 1, 255)
    End If
    
    Line Input #lNewf, sIn
    If Left(sIn, 6) = "!Dates" Then
        sDates = Replace(Mid(sIn, 8, 255), ",", " - ")
    End If
    
    grdNew.Cell(flexcpText, 0, 1) = sHdr & vbCrLf & sDates
    
    Do
    
        Line Input #lNewf, sIn
        If Left(sIn, 5) = "!Item" Then
            
            sIn = Mid(sIn, InStr(sIn, ",") + 1)
            
            sGroup = Left(sIn, InStr(sIn, ",") - 1)
            sIn = Mid(sIn, InStr(sIn, ",") + 1, 255)
            
            sItem = Left(sIn, InStr(sIn, ",") - 1)
            sIn = Mid(sIn, InStr(sIn, ",") + 1, 255)
           
            sSize = Left(sIn, InStr(sIn, ",") - 1)
            sIn = Mid(sIn, InStr(sIn, ",") + 1, 255)
        
            sFullWt = Left(sIn, InStr(sIn, ",") - 1)
            sIn = Mid(sIn, InStr(sIn, ",") + 1, 255)
        
            sEmptyWt = Left(sIn, InStr(sIn, ",") - 1)
            sIn = Mid(sIn, InStr(sIn, ",") + 1, 255)
        
            sIssueUnits = Left(sIn, InStr(sIn, ",") - 1)
            sIn = Mid(sIn, InStr(sIn, ",") + 1, 255)
        
            sNotes = sIn
            
            Line Input #lNewf, sIn
            If Left(sIn, 4) = "!Bar" Then
                            
                sIn = Mid(sIn, InStr(sIn, ",") + 1)
                sBar = Left(sIn, InStr(sIn, ",") - 1)
                
                sIn = Mid(sIn, InStr(sIn, ",") + 1)
                sFull = Left(sIn, InStr(sIn, ",") - 1)
                
                sIn = Mid(sIn, InStr(sIn, ",") + 1)
                sOpen = Left(sIn, InStr(sIn, ",") - 1)
                
                sIn = Mid(sIn, InStr(sIn, ",") + 1)
                sWeight = sIn
                
                
            
                grdNew.AddItem vbTab & "Group: " & sGroup & vbCrLf & _
                                        "Item: " & sItem & " " & sSize & "  " & "Full: " & sFullWt & "  " & "Empty: " & sEmptyWt & "  " & "Issue Units: " & sIssueUnits & vbCrLf & _
                                        "Bar: " & sBar & "  " & "Full: " & sFull & " " & "Open: " & sOpen & " " & "Weight: " & sWeight & vbCrLf & _
                                        "Note: " & sNotes
            Else
                sCount = Replace(Mid(sIn, InStr(sIn, ",") + 1, 255), ",", "")
            
                grdNew.AddItem vbTab & "Group: " & sGroup & vbCrLf & _
                                        "Item: " & sItem & " " & sSize & "  " & "Full : " & sFullWt & "  " & "Empty: " & sEmptyWt & "  " & "Issue Units: " & sIssueUnits & vbCrLf & _
                                        "Full: " & sFull & " " & "Open: " & sOpen & " " & "Weight: " & sWeight & vbCrLf & _
                                        "Note: " & sNotes
            
            End If
        
        End If
    
    Loop
    

Leave:
    Close #lNewf
    
    For i = 1 To grdNew.Rows - 1
    
        grdNew.RowHeight(i) = 1000
    Next
    
'    grdNew.Height = ((grdNew.Rows - 1) * 1000) + grdNew.RowHeightMin + 30
'    Me.Height = grdNew.Height + 1400
    
'    cmdQuit.Top = grdNew.Top + grdNew.Height + 150
'    cmdOk.Top = grdNew.Top + grdNew.Height + 150

    
    Exit Function

ErrorHandler:
    Resume Leave

    
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

Private Sub grdNew_Click()

    SetDeleteNewItemsFile

End Sub

Public Sub SetDeleteNewItemsFile()
Dim iRow As Integer

    For iRow = 1 To grdNew.Rows - 1
    
    
        If Not grdNew.Cell(flexcpChecked, iRow, 0) = 1 Then
            bDeleteNewItemsFile = False
            Exit Sub
        End If
    
    Next

    bDeleteNewItemsFile = True

End Sub
