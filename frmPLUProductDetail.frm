VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmPLUProductDetail 
   BackColor       =   &H00DCD5BC&
   BorderStyle     =   0  'None
   Caption         =   "StockWatch PLU / Product Detail"
   ClientHeight    =   8910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8940
   Icon            =   "frmPLUProductDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPLUProductDetail.frx":1CCA
   ScaleHeight     =   8910
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkHistory 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEF1E2&
      Caption         =   "Include in History Report"
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
      Height          =   375
      Left            =   4950
      TabIndex        =   32
      Top             =   795
      Value           =   1  'Checked
      Width           =   2490
   End
   Begin VSFlex8LCtl.VSFlexGrid grdSearch 
      Height          =   4230
      Left            =   2175
      TabIndex        =   21
      Top             =   4140
      Visible         =   0   'False
      Width           =   5445
      _cx             =   9604
      _cy             =   7461
      Appearance      =   2
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
      BackColorBkg    =   16777215
      BackColorAlternate=   14737632
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPLUProductDetail.frx":A288
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
      BackColorFrozen =   16711680
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.TextBox txtGlassDP 
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
      Left            =   7695
      MaxLength       =   8
      TabIndex        =   15
      Top             =   2910
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txtGlass 
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
      Left            =   3345
      MaxLength       =   8
      TabIndex        =   11
      Top             =   2910
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.TextBox txtSellDP 
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
      Left            =   5580
      MaxLength       =   8
      TabIndex        =   13
      Top             =   2910
      Visible         =   0   'False
      Width           =   795
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   330
      Top             =   5010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   596
      ImageHeight     =   593
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPLUProductDetail.frx":A2EE
            Key             =   "big"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPLUProductDetail.frx":128BC
            Key             =   "small"
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkActive 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEF1E2&
      Caption         =   "Active"
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
      Height          =   345
      Left            =   7695
      TabIndex        =   33
      Top             =   810
      Value           =   1  'Checked
      Width           =   885
   End
   Begin VB.TextBox txtPLU 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1380
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1155
      Width           =   795
   End
   Begin VB.TextBox txtSell 
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
      Left            =   1380
      MaxLength       =   8
      TabIndex        =   9
      Top             =   2910
      Width           =   780
   End
   Begin VB.ComboBox cboGroup 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1710
      Width           =   5445
   End
   Begin VB.TextBox txtPLUItem 
      BackColor       =   &H00DCD5BC&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   360
      Left            =   1380
      MaxLength       =   30
      TabIndex        =   6
      Top             =   2310
      Width           =   795
   End
   Begin VB.TextBox txtGroup 
      BackColor       =   &H00DCD5BC&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   360
      Left            =   1380
      MaxLength       =   30
      TabIndex        =   3
      Top             =   1710
      Width           =   795
   End
   Begin VB.TextBox txtCost 
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
      Left            =   1380
      MaxLength       =   10
      TabIndex        =   20
      Top             =   4350
      Width           =   795
   End
   Begin VB.TextBox txtStock 
      BackColor       =   &H00DCD5BC&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   360
      Left            =   1380
      MaxLength       =   20
      TabIndex        =   17
      Top             =   3780
      Width           =   795
   End
   Begin MyCommandButton.MyButton cmdSave 
      Height          =   495
      Left            =   7800
      TabIndex        =   24
      Top             =   5190
      Width           =   780
      _ExtentX        =   1376
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
      TransparentColor=   14472636
      Caption         =   "&Save"
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
   Begin MyCommandButton.MyButton cmdProductQuickAdd 
      Height          =   375
      Left            =   7620
      TabIndex        =   25
      Top             =   3750
      Width           =   960
      _ExtentX        =   1693
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
      TransparentColor=   14472636
      Caption         =   "Q Add"
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
   Begin MyCommandButton.MyButton cmdPLUQuickAdd 
      Height          =   375
      Left            =   7620
      TabIndex        =   26
      Top             =   2310
      Width           =   960
      _ExtentX        =   1693
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
      TransparentColor=   14472636
      Caption         =   "Q Add"
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
   Begin MyCommandButton.MyButton cmdSaveNew 
      Height          =   495
      Left            =   6390
      TabIndex        =   27
      Top             =   5190
      Width           =   1320
      _ExtentX        =   2328
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
      TransparentColor=   14472636
      Caption         =   "Save && &New"
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
      Left            =   5550
      TabIndex        =   28
      Top             =   5190
      Width           =   780
      _ExtentX        =   1376
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
      TransparentColor=   14472636
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
   Begin MyCommandButton.MyButton cmdDelete 
      Height          =   495
      Left            =   3570
      TabIndex        =   29
      Top             =   5190
      Width           =   1110
      _ExtentX        =   1958
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
      TransparentColor=   14472636
      Caption         =   "De&lete"
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
      Left            =   8520
      TabIndex        =   31
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
      TransparentColor=   14472636
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
   Begin VB.Label lblGlassDP 
      BackStyle       =   0  'Transparent
      Caption         =   "Glass Price 2"
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
      Left            =   6495
      TabIndex        =   14
      Top             =   2985
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblGlass 
      BackStyle       =   0  'Transparent
      Caption         =   "&Glass Price"
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
      Left            =   2280
      TabIndex        =   10
      Top             =   2970
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Label lblSellDP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sell Price 2"
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
      Left            =   4530
      TabIndex        =   12
      Top             =   2970
      Visible         =   0   'False
      Width           =   1005
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
      Height          =   465
      Left            =   1320
      TabIndex        =   30
      Top             =   1650
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label labelNoProds 
      Alignment       =   2  'Center
      BackColor       =   &H00DCD5BC&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4680
      TabIndex        =   23
      Top             =   1275
      Width           =   435
   End
   Begin VB.Label labelNoOfProducts 
      BackStyle       =   0  'Transparent
      Caption         =   "There are            Products Linked to this PLU#"
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
      Left            =   3540
      TabIndex        =   22
      Top             =   1305
      Visible         =   0   'False
      Width           =   4035
   End
   Begin VB.Label lblPLU 
      BackStyle       =   0  'Transparent
      Caption         =   "&PLU#"
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
      Left            =   840
      TabIndex        =   0
      Top             =   1215
      Width           =   615
   End
   Begin VB.Label lblGroup 
      BackStyle       =   0  'Transparent
      Caption         =   "&Group"
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
      Left            =   780
      TabIndex        =   2
      Top             =   1740
      Width           =   825
   End
   Begin VB.Label labelPLUItem 
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
      Height          =   360
      Left            =   2160
      TabIndex        =   7
      Top             =   2310
      Width           =   5445
   End
   Begin VB.Label lblSell 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sell P&rice"
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
      Left            =   300
      TabIndex        =   8
      Top             =   2970
      Width           =   1035
   End
   Begin VB.Label lblPLUItem 
      BackStyle       =   0  'Transparent
      Caption         =   "PL&U Item"
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
      Left            =   510
      TabIndex        =   5
      Top             =   2370
      Width           =   1125
   End
   Begin VB.Label lblCost 
      BackStyle       =   0  'Transparent
      Caption         =   "&Cost Price"
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
      Left            =   405
      TabIndex        =   19
      Top             =   4395
      Width           =   1245
   End
   Begin VB.Label labelStockItem 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   2160
      TabIndex        =   18
      Top             =   3780
      Width           =   5445
   End
   Begin VB.Label lblStock 
      BackStyle       =   0  'Transparent
      Caption         =   "Pr&oduct"
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
      Height          =   375
      Left            =   615
      TabIndex        =   16
      Top             =   3810
      Width           =   1335
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmPLUProductDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lPLUID As Long
Public lPLUGroupID As Long
Public lProductID As Long
Public lPLUProductID As Long
Public bNewProdPLUInProgress As Boolean
Private sSavedPLUNumber As String
Private iSavedNoOfOtherProductsLinked As Integer
Private bLoading As Boolean

Private Sub btnClose_Click()
    cmdQuit_Click

End Sub

Private Sub cboGroup_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub cboGroup_LostFocus()
    lblGroup.ForeColor = sBlack

'-------------------------------------------------------
' ver 530
      SetMeasure cboGroup
'-------------------------------------------------------
    
    If Trim$(Left(cboGroup, 2)) <> "*" Then
        txtGroup.BackColor = Me.BackColor
        txtGroup.ForeColor = Me.BackColor
    End If
    

End Sub

Private Sub chkActive_Click()

    If Not bLoading Then
    
        If chkActive = False Then
    
            MsgBox "Note: If this product has stock movement since last stock-take; Do not make inactive", vbOKOnly, "Deactivating Product"
        End If
    End If
    
End Sub

Private Sub cmdDelete_Click()
Dim iTopRow As Integer

    cmdDelete.Enabled = False
    cmdSave.Enabled = False
    cmdSaveNew.Enabled = False
    
    ' confirm delete
    If MsgBox("Are you sure you want to Delete This product?", vbDefaultButton2 + vbYesNo + vbQuestion, "Delete Product") = vbYes Then
    
        If DeleteProduct(lPLUProductID) Then
            LogMsg frmStockWatch, "Product " & labelStockItem & " - Deleted", "PLU#: " & txtPLU & " Cost: " & txtCost
        End If
        
    End If
    
    
    bHourGlass False
    
'    cmdDelete.Enabled = True
'    cmdSave.Enabled = True
'    cmdSaveNew.Enabled = True
        
    iTopRow = frmCtrl.grdList.TopRow
        
    gbOk = frmCtrl.ShowProductPLUs(frmCtrl.cboActive.ListIndex, frmCtrl.cboByGroup.ListIndex, frmCtrl.txtSearch)
        
    frmCtrl.grdList.TopRow = iTopRow
    
    cmdQuit_Click
    
    ' Warn removing PLU also if only one product pointing to plu
    

End Sub

Private Sub cmdPLUQuickAdd_Click()

    frmPLUDetail.lPLUID = 0 ' make sure since this is definitely an add new request
    frmPLUDetail.bQuickAdd = True
    frmPLUDetail.Show vbModal
    grdSearch.Visible = False

    lPLUID = frmPLUDetail.lPLUID
    labelPLUItem = GetPLUDescription(lPLUID)
    bSetFocus Me, "txtSell"
    
    ' here we display the returned PLU and save the PLUID


End Sub

Private Sub cmdProductQuickAdd_Click()
    
    frmProductDetail.lProductID = 0 ' make sure since this is definitely an add new request
    frmProductDetail.bQuickAdd = True
    frmProductDetail.Show vbModal
    grdSearch.Visible = False

    lProductID = frmProductDetail.lProductID
    labelStockItem = GetProductDescription(lProductID)
    bSetFocus Me, "txtCost"
    
    ' here we display the returned PLU and save the PLUID



End Sub

Private Sub cmdQuit_Click()


    Unload Me
    

End Sub

Private Sub cmdSave_Click()
Dim iSavedRow As Integer

    
    If frmCtrl.grdList.Row <> frmCtrl.grdList.Rows - 1 Then
        iSavedRow = frmCtrl.grdList.Row
    End If
    ' Save the current position on the grid
    ' rev 34
    
    If DoSave() Then
        
        frmCtrl.bAddNewPLUProduct = True
        
        frmCtrl.grdList.Row = iSavedRow + 1
        ' rev 304
        
        cmdQuit_Click
    
    End If
    
    bHourGlass False
    
End Sub

Private Sub cmdSave_GotFocus()

    SetFormSmall True

End Sub

Private Sub cmdSaveNew_Click()

    If DoSave() Then

        frmCtrl.bAddNewPLUProduct = True
        
        InitPLUItem

        lPLUProductID = 0
        
        txtPLU.Text = ShowNextPLUNumber(lSelClientID)
    
        bSetFocus Me, "txtPLU"
    End If
    
End Sub

Private Sub cmdSaveNew_GotFocus()

    SetFormSmall True

End Sub

Private Sub Form_Activate()

    bHourGlass False
    
    If lPLUProductID <> 0 Then
        
'-------------------------------------------------
' ver 554 as PermissionEnum Kate email request

        bSetFocus Me, "txtPLU"
        
'        If Val(txtSell) = 0 Then
'            bSetFocus Me, "txtSell"
'        ElseIf txtGlass.Visible And Val(txtGlass) = 0 Then
'            bSetFocus Me, "txtGlass"
'        ElseIf bDualPrice And Val(txtSellDP.Text) = 0 Then
'            bSetFocus Me, "txtSellDP"
'        ElseIf txtGlassDP.Visible And Val(txtGlassDP.Text) = 0 Then
'            bSetFocus Me, "txtGlassDP"
'        End If
'-------------------------------------------------
    
    End If
    
    Me.Top = Screen.Height / 4
    
    bLoading = False
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    ' return pressed force focus to next available object in tabbing order
        gbOk = GotoNextControl(Me, 0)
    
    
    ElseIf KeyAscii = 27 Then
        
        
        If grdSearch.Visible Then
            grdSearch.Visible = False
            SetFormSmall True
        
        Else
        
            cmdQuit_Click
        End If
        
    End If



End Sub

Private Sub Form_Load()
         
    bLoading = True
         
    gbOk = GetPLUGroupList(lSelClientID)
    ' get the list of PLU groups for this client
    ' if they're not there already show generic list

    InitPLUItem
    
    If lPLUProductID <> 0 Then
    ' will be some value for an edit
    
        gbOk = ShowPLUProduct(lPLUProductID)
        
'-------------------------------------------------------
' ver 530
                SetMeasure cboGroup
'-------------------------------------------------------
        
        sSavedPLUNumber = txtPLU.Text
        gbOk = GetNoOfProductsForPLU(txtPLU)
        
        iSavedNoOfOtherProductsLinked = Val(labelNoProds)
        
        bSetFocus Me, "txtSell"
    
    Else
        txtPLU.Text = ShowNextPLUNumber(lSelClientID)
    End If

    ' ver 3.0.6 (6) removed setok routine here because franchisee
    ' should be able to add/moddify any product
    

End Sub

Private Sub grdSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        If grdSearch.Row = 0 Then
    
            Select Case grdSearch.Tag
    
                Case "PLU"
                 bSetFocus Me, "txtPLUItem"
                Case Else
                 bSetFocus Me, "txtStock"
            
            End Select
        
        End If
    
    End If
    

End Sub

Private Sub grdSearch_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then grdSearch_Click

End Sub

Private Sub Label1_Click()

End Sub

Private Sub txtCost_GotFocus()
    
    SetFormSmall True
    
    gbOk = bSetupControl(Me)
    
    txtStock.Width = txtCost.Width ' just to be sure
    
'    grdSearch.Visible = False

End Sub

Private Sub txtCost_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    
        If lPLUProductID = 0 Then
            bSetFocus Me, "cmdSaveNew"
        Else
            bSetFocus Me, "cmdSave"
        End If
        
    ElseIf KeyAscii <> 46 Then
        
        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
        
    ElseIf InStr(txtCost, ".") > 0 Then
            
        KeyAscii = 0
    End If
    

End Sub

Private Sub txtCost_LostFocus()
    lblCost.ForeColor = sBlack
    

End Sub

Private Sub txtGlass_GotFocus()
    SetFormSmall True
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtGlass_LostFocus()
    lblGlass.ForeColor = sBlack

End Sub

Private Sub txtGlassDP_GotFocus()
    SetFormSmall True
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtGlassDP_LostFocus()
    lblGlassDP.ForeColor = sBlack

End Sub

Private Sub txtGroup_GotFocus()
    
    SetFormSmall True
    
    gbOk = bSetupControl(Me)
    txtGroup.BackColor = sWhite
    txtGroup.ForeColor = sBlack

End Sub

Private Sub txtGroup_KeyPress(KeyAscii As Integer)
    KeyAscii = CharOk(KeyAscii, 0, "") ' 0 = no only, 1 = char only, 2 = both
    
End Sub

Private Sub txtGroup_LostFocus()

    If txtGroup.Text = "" Then
        txtGroup.BackColor = Me.BackColor
    End If
    

End Sub

Private Sub txtPLU_GotFocus()
    
    SetFormSmall True
    
    gbOk = bSetupControl(Me)
'    grdSearch.Visible = False
    ' just make sure the search panel is hidden

End Sub

Private Sub txtPLU_KeyPress(KeyAscii As Integer)
    KeyAscii = CharOk(KeyAscii, 0, "") ' 0 = no only, 1 = char only, 2 = both

End Sub

Private Sub txtPLU_LostFocus()
Dim lID As Long
    
    ' show plu if it exists
    '   then move plu item to txtstock and set focus there
    ' otherwise
    '   goto txtgroup
    
    lblPLU.ForeColor = sBlack

    If lPLUID = 0 Or (sSavedPLUNumber <> txtPLU) Then
    
        If ShowPLU(txtPLU, lPLUID) Then
        
            gbOk = GetNoOfProductsForPLU(txtPLU)
        
            bSetFocus Me, "txtStock"
            
            txtStock_KeyUp 13, 0
        
        Else
            bSetFocus Me, "txtGroup"
        End If
    
' ver 555-----------------------
    Else
        bSetFocus Me, "txtSell"
    End If
    
'    ElseIf Val(txtSell) = 0 Then
'        bSetFocus Me, "txtSell"
'    ElseIf txtGlass.Visible And Val(txtGlass) = 0 Then
'        bSetFocus Me, "txtGlass"
'    ElseIf bDualPrice And Val(txtSellDP.Text) = 0 Then
'        bSetFocus Me, "txtSellDP"
'    ElseIf txtGlassDP.Visible And Val(txtGlassDP.Text) = 0 Then
'        bSetFocus Me, "txtGlassDP"
'    End If

' ver 555-----------------------
    
End Sub

Public Function ShowPLU(sPLU As String, lID As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    If sPLU <> "" Then
        Set rs = SWdb.OpenRecordset("SELECT PLUID, PLUGroupID, txtDescription, SellPrice, Active, SellPriceDP FROM tblClientProductPLUs INNER JOIN tblPLUs ON tblClientProductPLUs.PLUID = tblPLUs.ID WHERE PLUNumber = " & sPLU & " AND CLientID = " & Trim$(lSelClientID), dbOpenSnapshot)
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
            gbOk = PointToEntry(Me, "cboGroup", rs("PLUGroupID"), False)
            labelPLUItem.Caption = rs("txtDescription")
            txtSell.Text = Format(rs("SellPrice"), "0.00")
            txtSellDP.Text = Format(rs("SellPriceDP"), "0.00")
            chkActive.Value = Abs(rs("Active"))
            
            lPLUID = rs("PLUID")
            lPLUGroupID = rs("PLUGroupID")
        
            ShowPLU = True
        
        Else
            ShowPLU = False
    
        End If
        rs.Close
    End If
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("ShowPLU") Then Resume 0
    Resume CleanExit


End Function

Private Sub txtPLUItem_GotFocus()
    
    SetFormSmall True
    
    gbOk = bSetupControl(Me)
    txtPLUItem.BackColor = sWhite
    txtPLUItem.ForeColor = sBlack

End Sub

Private Sub txtPLUItem_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If txtPLUItem <> "" And (lPLUID = 0 Or labelPLUItem = "") Then
            grdSearch.Visible = True
            SetFormSmall False

            If grdSearch.Rows > 1 Then grdSearch.Row = 0
            bSetFocus Me, "grdSearch"
        Else
            grdSearch.Visible = False
            SetFormSmall True
        
        End If
    Else
        KeyAscii = CharOk(KeyAscii, 2, " /*'&") ' 0 = no only, 1 = char only, 2 = both

    End If
    

End Sub

Private Sub txtPLUItem_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Then
    
    ElseIf Len(txtPLUItem) > 0 Then
    
        If KeyCode = vbKeyDown Then
        
            If grdSearch.Visible Then
                bSetFocus Me, "grdSearch"
                grdSearch.Row = 0
            End If
        
        Else
            grdSearch.Tag = "PLU"
            gbOk = SearchPLUList(txtPLUItem)
    
        End If
    
    ElseIf KeyCode = 8 Then
        grdSearch.Visible = False
    End If
    
End Sub

Private Sub txtPLUItem_LostFocus()
    txtPLUItem.BackColor = Me.BackColor
    lblPLUItem.ForeColor = sBlack
    txtPLUItem.ForeColor = Me.BackColor
    
End Sub

Private Sub txtSell_GotFocus()
            
    SetFormSmall True
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtSell_KeyPress(KeyAscii As Integer)
    
        
    If KeyAscii <> 46 Then
        
        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
        
    ElseIf InStr(txtSell, ".") > 0 Then
            
        KeyAscii = 0
        
    End If

End Sub

Private Sub txtSell_LostFocus()
    lblSell.ForeColor = sBlack
    
End Sub

Private Sub txtSellDP_GotFocus()
    SetFormSmall True
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtSellDP_KeyPress(KeyAscii As Integer)
        
    If KeyAscii <> 46 Then
        
        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
        
    ElseIf InStr(txtSellDP, ".") > 0 Then
            
        KeyAscii = 0
        
    End If

End Sub

Private Sub txtSellDP_LostFocus()
    lblSellDP.ForeColor = sBlack

End Sub

Private Sub txtStock_GotFocus()
    
    SetFormSmall True
    
    txtStock.BackColor = sWhite
    txtStock.ForeColor = sBlack
    
    If bNewProdPLUInProgress Then
    
        txtStock.Text = SuggestStockItem(labelPLUItem.Caption)
        
        SetBoxWidth
        
    End If
    
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtStock_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If txtStock <> "" And (lProductID = 0 Or labelStockItem = "") Then
            grdSearch.Visible = True
            If grdSearch.Rows > 0 Then grdSearch.Row = 0
            bSetFocus Me, "grdSearch"
            SetFormSmall False

        Else
            grdSearch.Visible = False
            SetFormSmall True
        
        End If
    
    Else
            KeyAscii = CharOk(KeyAscii, 2, " /'*&") ' 0 = no only, 1 = char only, 2 = both
            SetBoxWidth

    End If

End Sub

Private Sub txtStock_LostFocus()
    txtStock.BackColor = Me.BackColor
    lblStock.ForeColor = sBlack
    txtStock.ForeColor = Me.BackColor
    txtStock.Width = txtCost.Width

End Sub
Private Sub txtGroup_KeyUp(KeyCode As Integer, Shift As Integer)
Dim iCnt As Integer

    ' here we're trying to find a match in the Group drop down list box
    ' with the group number just entered
    ' Then save the ID
    
    If Val(txtGroup) > 0 Then
    ' if a group number is entered.... check for it on the list....
    
        For iCnt = 0 To cboGroup.ListCount - 1
            If Val(cboGroup.List(iCnt)) = Val(txtGroup) Then
            ' group found so point to it on the list
            
                cboGroup.ListIndex = iCnt
'                lPLUGroupID = cboGroup.ItemData(cboGroup.ListIndex)
'                ' Also Save the PLU Group ID
            
                
'-------------------------------------------------------
' ver 530
                SetMeasure cboGroup
'-------------------------------------------------------
                
                
                
                Exit For
            End If
            
        Next
    End If
    
End Sub

Private Sub txtStock_KeyUp(KeyCode As Integer, Shift As Integer)
Dim sItem As String

    If KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Then
    
        If grdSearch.Visible Then
            bSetFocus Me, "grdSearch"
            If grdSearch.Rows > 0 Then grdSearch.Row = 0
        End If
        
    ElseIf Val(txtStock) > 999 Then
       If GetSTKItem(Val(txtStock)) Then
            grdSearch.Visible = False
       End If
       
    ElseIf KeyCode = 27 Then
    
    
    ElseIf Len(txtStock) > 0 Then
       grdSearch.Tag = ""
       
       gbOk = SearchSTKList(txtStock)
    
    ElseIf KeyCode = 8 Then
        grdSearch.Visible = False

    End If

End Sub

Public Function GetPLUGroupList(lCLId As Long)
Dim rs As Recordset
Dim iCnt As Integer

    On Error GoTo ErrorHandler
    
    ' show list like so:
    
    '       1 Bottled Beers
    '       2 Draught Beers
    '       * Spirits
    '       * Wines
    '   etc...
    
    cboGroup.Clear
    
    Set rs = SWdb.OpenRecordset("SELECT * from tblPLUGroup WHERE ClientID = " & Trim$(lCLId) & " ORDER BY txtGroupNumber", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        Do
            cboGroup.AddItem Trim$(rs("txtGroupNumber")) & "   " & rs("txtDescription")
            cboGroup.ItemData(cboGroup.NewIndex) = rs("ID") + 0
            rs.MoveNext
        Loop While Not rs.EOF
        
    End If
    ' These are the Groups that are already numbered for this client
    
            
    Set rs = SWdb.OpenRecordset("SELECT DISTINCT txtDescription from tblPLUGroup ORDER BY txtDescription", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        Do
            For iCnt = 0 To cboGroup.ListCount - 1
                If InStr(cboGroup.List(iCnt), rs("txtDescription")) <> 0 Then
                    ' found it...
                    GoTo GetNext
                End If
            Next
    
            cboGroup.AddItem "*   " & rs("txtDescription")
            
GetNext:
        rs.MoveNext
        Loop While Not rs.EOF
    
    End If
    ' These are the other groups that havent been numbered yet
    ' An astrix is shown in the numbers place
    
    
    GetPLUGroupList = True
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
'    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetPLUGroupList") Then Resume 0
    Resume CleanExit



End Function

Public Function GetSTKItem(iCode As Integer)
Dim rs As Recordset
    
    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("Select * FROM tblProducts INNER JOIN tblproductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE txtCode = '" & Trim$(iCode) & "'", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        labelStockItem = rs("txtCode") & "  " & rs("tblProductGroup.txtDescription") & "  " & rs("tblProducts.txtDescription") & "    " & rs("txtsize")
        
        lProductID = rs("tblProducts.ID") + 0
        ' Save Product ID here
        
        grdSearch.Visible = False
    
    End If
    
    GetSTKItem = True
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
'
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetSTKItem") Then Resume 0
    Resume CleanExit

End Function

Public Function SearchSTKList(sProduct As String)
Dim rs As Recordset
    
    On Error GoTo ErrorHandler
    
    grdSearch.Rows = 0
    grdSearch.Top = labelStockItem.Top + labelStockItem.Height + 20
    grdSearch.Left = labelStockItem.Left
    
    grdSearch.ColHidden(0) = False
    grdSearch.ColHidden(1) = False
    grdSearch.ColWidth(1) = 3350
    grdSearch.Left = txtStock.Left
    grdSearch.Width = labelStockItem.Width + txtStock.Width
    grdSearch.ColHidden(2) = False
    grdSearch.Tag = "STK"
    
    ' ver 3.0.1 added chkactive flag check
    
    Set rs = SWdb.OpenRecordset("Select * FROM tblProducts INNER JOIN tblproductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE tblPRoducts.chkActive = true AND tblProducts.txtDescription LIKE " & """*" & sProduct & "*""", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do
        
            grdSearch.AddItem rs("tblProductGroup.txtDescription") & vbTab & rs("tblProducts.txtDescription") & vbTab & rs("txtsize")
            grdSearch.RowData(grdSearch.Rows - 1) = rs("tblProducts.ID") + 0
        
            rs.MoveNext
        Loop While Not rs.EOF
    
        SetFormSmall False
        
        grdSearch.Visible = True
    
    End If

    SearchSTKList = True
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
'
    
    Exit Function

ErrorHandler:
    If CheckDBError("SearchSTKList") Then Resume 0
    Resume CleanExit

End Function
Private Sub grdSearch_LostFocus()

    SetFormSmall True
End Sub

Private Sub grdSearch_Click()

    If grdSearch.Row > -1 And grdSearch.Row < grdSearch.Rows Then
    
        Select Case grdSearch.Tag
        
            Case "PLU"
             labelPLUItem.Caption = grdSearch.Cell(flexcpTextDisplay, grdSearch.Row, 1)
             lPLUID = grdSearch.RowData(grdSearch.Row)
             ' Save PLU ID here
             
             bSetFocus Me, "txtSell"
            
            Case Else
             labelStockItem.Caption = grdSearch.Cell(flexcpTextDisplay, grdSearch.Row, 0) & " " & grdSearch.Cell(flexcpTextDisplay, grdSearch.Row, 1) & " " & grdSearch.Cell(flexcpTextDisplay, grdSearch.Row, 2)
             
             lProductID = grdSearch.RowData(grdSearch.Row)
             ' Save Product ID here
    
             bSetFocus Me, "txtCost"
        
        End Select
    
        grdSearch.Visible = False
        
        txtPLUItem.BackColor = Me.BackColor
        txtStock.BackColor = Me.BackColor
    
        lblPLUItem.ForeColor = sBlack
        lblStock.ForeColor = sBlack
    
    End If

End Sub
Public Function SearchPLUList(sPLU As String)
Dim rs As Recordset
    
    On Error GoTo ErrorHandler
    
    grdSearch.Rows = 0
    
    grdSearch.Top = labelPLUItem.Top + labelPLUItem.Height + 20
    grdSearch.Left = labelPLUItem.Left + 20
    grdSearch.Width = labelStockItem.Width
    
    grdSearch.ColHidden(0) = True
    grdSearch.ColHidden(1) = False
    grdSearch.ColWidth(1) = 5200

    grdSearch.ColHidden(2) = True
    
    grdSearch.Tag = "PLU"
    
    ' Ver 3.0.9 include active shieck
    Set rs = SWdb.OpenRecordset("Select * FROM tblPLUs WHERE chkActive = true AND txtDescription LIKE " & """" & sPLU & "*""" & " ORDER BY txtDescription", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do
        
            grdSearch.AddItem vbTab & rs("txtDescription")
            grdSearch.RowData(grdSearch.Rows - 1) = rs("ID") + 0
        
            rs.MoveNext
        Loop While Not rs.EOF
    
        SetFormSmall False
        grdSearch.Visible = True
    
    End If
    
    SearchPLUList = True
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
'
    
    Exit Function

ErrorHandler:
    If CheckDBError("SearchPLUList") Then Resume 0
    Resume CleanExit


End Function

Public Function SavePLU(lPLUPRodID As Long, dblSalesQty As Double, dblSalesQtyDP As Double, dblGlassQty As Double, dblGlassQtyDP As Double)
Dim rs As Recordset
    
' ver530 adding glasses here

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblClientProductPLUs")
    rs.Index = "PrimaryKey"
    If lPLUPRodID > 0 Then
        rs.Seek "=", lPLUPRodID
        If Not rs.NoMatch Then
            rs.Edit
        Else
            rs.AddNew
            rs("FullQty") = Empty
            rs("Open") = Empty
            rs("weight") = Empty
            rs("RcvdQty") = Empty
            rs("SalesQty") = dblSalesQty
            rs("SalesQtyDP") = dblSalesQtyDP
            rs("GlassQty") = dblGlassQty
            rs("GlassQtyDP") = dblGlassQtyDP

        End If
    Else
        rs.AddNew
        rs("FullQty") = Empty
        rs("Open") = Empty
        rs("weight") = Empty
        rs("RcvdQty") = Empty
        rs("SalesQty") = dblSalesQty
        rs("SalesQtyDP") = dblSalesQtyDP
        rs("GlassQty") = dblGlassQty
        rs("GlassQtyDP") = dblGlassQtyDP
    
    End If
    
    rs("ClientID") = lSelClientID
    rs("PLUNumber") = Val(txtPLU)
    
    rs("PLUID") = lPLUID
    rs("PLUGroupID") = lPLUGroupID
    rs("SellPrice") = Format(txtSell, "0.00")
    
    'DP
    If Val(txtSellDP) <> 0 Then
        rs("SellPriceDP") = Format(txtSellDP, "0.00")
    End If
    
    rs("ProductID") = lProductID
    rs("PurchasePrice") = txtCost
    
    rs("Active") = chkActive
    
    'Ver550
    rs("chkHistory") = chkHistory
    '------
    rs.Update

    SavePLU = True
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
'
    
    Exit Function

ErrorHandler:
    
    If Err = 3022 Then
        rs("id") = rs("id") + 1
        Resume 0
    
    Else
    
        If CheckDBError("SavePLU") Then Resume 0
        Resume CleanExit
    End If

End Function
Public Function UpdatePriceOtherItemsSamePLU(lCLId As Long, sPLU As String, lPLUID As Long, lPLUGroupID, sSell As String, sSellDP As String, sGlass As String, sGlassDP As String)
Dim curSellDP As Currency
Dim curGlassDP As Currency

    curSellDP = Val(sSellDP)
    curGlassDP = Val(sGlassDP)
    
    If lCLId > 0 Then
    ' make sure we've passed some legitimate values
    
        'DP
'        SWdb.Execute "UPDATE tblClientProductPLUs SET PLUID = " & Trim$(lPLUID) & ", PLUGroupID = " & Trim$(lPLUGroupID) & ", SellPrice = " & Trim$(Format(sSell, "0.00")) & " WHERE ClientID = " & Trim$(lCLId) & " AND PLUNumber = " & sPLU
'        SWdb.Execute "UPDATE tblClientProductPLUs SET PLUID = " & Trim$(lPLUID) & ", PLUGroupID = " & Trim$(lPLUGroupID) & ", SellPrice = " & Trim$(Format(sSell, "0.00")) & ", SellPriceDP = " & curSellDP & " WHERE ClientID = " & Trim$(lCLId) & " AND PLUNumber = " & sPLU
        
' ver530
        SWdb.Execute "UPDATE tblClientProductPLUs SET PLUID = " & Trim$(lPLUID) & ", PLUGroupID = " & Trim$(lPLUGroupID) & ", SellPrice = " & Trim$(Format(Val(sSell), "0.00")) & ", SellPriceDP = " & Trim$(Format(Val(curSellDP), "0.00")) & ", GlassPrice = " & Trim$(Format(Val(sGlass), "0.00")) & ", GlassPriceDP = " & Trim$(Format(Val(curGlassDP), "0.00")) & " WHERE ClientID = " & Trim$(lCLId) & " AND PLUNumber = " & sPLU
        
        UpdatePriceOtherItemsSamePLU = True
        
    End If
    
End Function

Public Sub InitPLUItem()

    SetFormSmall True
    
    cboGroup.ListIndex = -1
    txtGroup.Text = ""
    txtPLUItem.Text = ""
    labelPLUItem.Caption = ""
    txtStock.Text = ""
    labelStockItem.Caption = ""
    txtCost.Text = ""
    txtSell.Text = ""
    txtSellDP.Text = ""
    txtGlass.Text = ""
    txtGlassDP.Text = ""
    chkActive.Value = 1
    chkHistory.Value = 1
    lPLUID = 0
    lProductID = 0
    lPLUGroupID = 0
    labelNoProds = ""
    labelNoOfProducts.Visible = False
    
    txtStock.Width = txtCost.Width

    bNewProdPLUInProgress = False
    sSavedPLUNumber = ""
    iSavedNoOfOtherProductsLinked = 0

        
    
    lblSellDP.Visible = bDualPrice
    txtSellDP.Visible = bDualPrice
    lblGlassDP.Visible = bDualPrice
    txtGlassDP.Visible = bDualPrice
    'DP

End Sub

Public Function FieldsCheck()

    If Val(txtPLU.Text) > 0 Then
        If (Val(txtGroup.Text) > 0) Or (Val(Left(cboGroup, InStr(1, cboGroup, " "))) > 0) Then
            If cboGroup.ListIndex > -1 Then
                If labelPLUItem.Caption <> "" Then
                    If labelStockItem.Caption <> "" Then
                        
                        If SellCostPriceOK(txtSell.Text, "Sell") Then
                            
                            If SellCostPriceOK(txtCost.Text, "Cost") Then
                                
                                FieldsCheck = True
    
                            Else
                                bSetFocus Me, "txtCost"
                            End If
                        Else
                            bSetFocus Me, "txtSell"
                        End If
                    Else
                        MsgBox "Please Select a Valid Stock Item from the list"
                        bSetFocus Me, "txtStock"
                    End If
                Else
                    MsgBox "Please Select a Valid PLU Description"
                    bSetFocus Me, "txtPLUItem"
                End If
            Else
                MsgBox "Please Select a PLU Group"
                bSetFocus Me, "cboGroup"
            End If
    
        Else
            MsgBox "Please Set a PLU group Number"
            bSetFocus Me, "txtGroup"
        End If
    Else
        MsgBox "Please enter a valid PLU Number"
        bSetFocus Me, "txtPLU"
    End If

End Function

Public Function SavePLUGroup(lCLId As Long, iGroupNumber As Integer, sGroupName As String)
Dim rs As Recordset
    
    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("SELECT * FROM tblPLUGroup WHERE ClientID = " & Trim$(lCLId) & " AND txtGroupNumber = " & Trim$(iGroupNumber))
    ' first make sure its not there already
    
    If (rs.EOF And rs.BOF) Then
    
        Set rs = SWdb.OpenRecordset("tblPLUGroup")
        rs.Index = "PrimaryKey"
        rs.AddNew
        rs("ClientID") = lCLId
        rs("txtGroupNumber") = iGroupNumber
        rs("txtDescription") = Trim$(Replace(sGroupName, "*", ""))
        rs("chkActive") = True
        rs.Update
        rs.Bookmark = rs.LastModified
        
    End If

    SavePLUGroup = rs("ID") + 0
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
'
    
    Exit Function

ErrorHandler:
    
    If Err = 3022 Then
        rs("id") = rs("id") + 1
        Resume 0
    
    Else
    
        If CheckDBError("SavePLUGroup") Then Resume 0
        Resume CleanExit
    End If
    
End Function

Public Function SellCostPriceOK(sAmount As String, sSellCostPrice As String)

    If Val(sAmount) = 0 Then
        If MsgBox("Are you sure you want to set the " & sSellCostPrice & " Price = 0 ?", vbDefaultButton1 + vbYesNo + vbQuestion, "Setting Price=0") = vbYes Then
            
            Select Case sSellCostPrice
                Case "Sell"
                 txtSell.Text = "0.00"
                Case "Cost"
                 txtCost.Text = "0.00"
            
            End Select
            
            SellCostPriceOK = True
        
        End If
    Else
        SellCostPriceOK = True
    
    End If

End Function

Public Function ShowPLUProduct(lPID As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("SELECT * FROM (((tblClientProductPLUs INNER JOIN tblPLUs ON tblClientProductPLUs.PLUID = tblPLUs.ID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID) INNER JOIN tblPLUGroup ON tblClientProductPLUs.PLUGroupID = tblPLUGroup.ID WHERE tblClientProductPLUs.ID = " & Trim$(lPID), dbOpenSnapshot)
'    Set rs = SWdb.OpenRecordset("SELECT * FROM (((tblClientProductPLUs LEFT JOIN tblPLUs ON tblClientProductPLUs.PLUID = tblPLUs.ID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID) INNER JOIN tblPLUGroup ON tblClientProductPLUs.PLUGroupID = tblPLUGroup.ID WHERE tblClientProductPLUs.ID = " & Trim$(lPID), dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        gbOk = PointToEntry(Me, "cboGroup", rs("tblPLUGroup.ID"), False)
        labelPLUItem.Caption = rs("tblPLUs.txtDescription")
        labelStockItem.Caption = rs("PLUNumber") & "   " & rs("tblProductGroup.txtDescription") & "    " & rs("tblProducts.txtDescription") & "    " & rs("txtsize")
        txtCost.Text = Format(rs("PurchasePrice"), "0.00")
        txtSell.Text = Format(rs("SellPrice"), "0.00")
        txtSellDP.Text = Format(rs("SellPriceDP"), "0.00")
        txtGlass.Text = Format(rs("glassPrice"), "0.00")
        txtGlassDP.Text = Format(rs("glassPriceDP"), "0.00")
        chkActive.Value = Abs(rs("Active"))
        
        'Ver550
        chkHistory.Value = Abs(rs("chkHistory"))
        '------

        
        txtPLU.Text = rs("PLUNumber")
        
        lPLUID = rs("tblPLUs.ID")
        lProductID = rs("tblProducts.ID")
        lPLUGroupID = rs("tblPLUGroup.ID")
        lPLUProductID = lPID
    
 
    End If
    
    ShowPLUProduct = True
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
'
    
    Exit Function

ErrorHandler:
    If CheckDBError("ShowPLUProduct") Then Resume 0
    Resume CleanExit

End Function


Public Function GetPLUDescription(lID As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    If lID > 0 Then
        
        Set rs = SWdb.OpenRecordset("tblPLUs")
        rs.Index = "PrimaryKey"
        rs.Seek "=", lID
        If Not rs.NoMatch Then
            GetPLUDescription = rs("txtDescription") & ""
        End If
    
        rs.Close
    
    Else
        GetPLUDescription = ""
    End If
    

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetPLUDescription") Then Resume 0
    Resume CleanExit

End Function

Public Function GetProductDescription(lID As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    If lID > 0 Then
        
        Set rs = SWdb.OpenRecordset("SELECT * FROM tblProducts INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE tblProducts.ID = " & Trim$(lID), dbOpenSnapshot)
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
            GetProductDescription = rs("txtCode") & "  " & rs("tblProductGroup.txtDescription") & "  " & rs("tblProducts.txtDescription") & "    " & rs("txtsize")
        End If
    
        rs.Close
    
    Else
        GetProductDescription = ""
    End If
    

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetPLUDescription") Then Resume 0
    Resume CleanExit

End Function


Public Function DoSave()
Dim iTopRow As Integer
Dim dblSalesQty As Double
Dim dblSalesQtyDP As Double
Dim dblGlassQty As Double
Dim dblGlassQtyDP As Double

    If FieldsCheck() Then
    
'Ver433 - small check to see if part present already

      If PartNotPresentAlready(lProductID, lSelClientID) Then
        
        cmdSave.Enabled = False
        cmdSaveNew.Enabled = False
        cmdDelete.Enabled = False
        ' PLU Group
        
        ' Save new plu group if required
        
        If Trim$(Left(cboGroup, 2)) = "*" Then
        ' No group defined so go set it up for this product
        
            lPLUGroupID = SavePLUGroup(lSelClientID, Val(txtGroup), cboGroup)
        Else
            lPLUGroupID = cboGroup.ItemData(cboGroup.ListIndex)
            
        End If
        
'Ver 433 =====================================================================
        
        gbOk = GetQTYSForThisPLU(lSelClientID, txtPLU, dblSalesQty, dblSalesQtyDP, dblGlassQty, dblGlassQtyDP)
        
        gbOk = SavePLU(lPLUProductID, dblSalesQty, dblSalesQtyDP, dblGlassQty, dblGlassQtyDP)
         
        LogMsg frmStockWatch, "Product/PLU Saved for " & Replace(frmStockWatch.lblClient.Tag, "_", " "), " PLU#:" & txtPLU & " Group#:" & txtGroup & " Group:" & cboGroup & " PLU Item:" & labelPLUItem & " Price:" & txtSell & " Stock Item:" & labelStockItem & " Cost:" & txtCost & " Act:" & Trim$(chkActive) & " His:" & Trim$(chkHistory)
        
        'DP
        gbOk = UpdatePriceOtherItemsSamePLU(lSelClientID, txtPLU, lPLUID, lPLUGroupID, txtSell, txtSellDP, txtGlass, txtGlassDP)
        ' in case the there are other items linked to the same PLU we
        ' must make sure the same price appears on all these items.
         

        If lPLUProductID <> 0 Then
        ' an edit ... so check is there more than one product tied to this plu
        ' and then ... was it the plu number that was changed?
        
            If iSavedNoOfOtherProductsLinked > 1 Then
            ' more than one product linked....
            
                If txtPLU.Text <> sSavedPLUNumber Then
                ' plu# was changed...
                
                    If MsgBox("Move " & Trim$(iSavedNoOfOtherProductsLinked - 1) & " other Product(s) to this PLU Number?", vbYesNo + vbQuestion + vbDefaultButton2, "PLU Number Changed") = vbYes Then
        
                        gbOk = UpdatePLUNumberOnOtherProducts(lSelClientID, sSavedPLUNumber, txtPLU)
                    
                    End If
                End If
            End If
            
        
        End If
        
        gbOk = GetPLUGroupList(lSelClientID)
        ' get the list of PLU groups for this client
        ' if they're not there already show generic list
         
        txtPLU.Text = ""
         
        InitPLUItem
         
'ver304
        'cmdSave.Enabled = True
        'cmdSaveNew.Enabled = True
        
        iTopRow = frmCtrl.grdList.TopRow
        
        gbOk = frmCtrl.ShowProductPLUs(frmCtrl.cboActive.ListIndex, frmCtrl.cboByGroup.ListIndex, frmCtrl.txtSearch)
        
        frmCtrl.grdList.TopRow = iTopRow
        '

        cmdSave.Enabled = True
        cmdSaveNew.Enabled = True
        cmdDelete.Enabled = True
        ' PLU Group
        
        DoSave = True
    
    
      Else
        MsgBox "This product is already on the list"
              
      End If
    
    End If

End Function

Public Function ShowNextPLUNumber(lID As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler

    Set rs = SWdb.OpenRecordset("SELECT Max(PLUNumber) as LastPLU FROM tblClientProductPLUs WHERE ClientID = " & Trim$(lID), dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        If Not IsNull(rs("LastPLU")) Then
            ShowNextPLUNumber = rs("LastPLU") + 1
        Else
            ShowNextPLUNumber = 1
        End If
        
    End If
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("ShowNextPLUNumber") Then Resume 0
    Resume CleanExit

End Function
Public Function SuggestStockItem(sItem As String)
Dim sStock As String
    
    sStock = Trim$(Replace(sItem, "1/2", ""))
    sStock = Trim$(Replace(sItem, "1/4", ""))
    sStock = Trim$(Replace(sStock, "BOT ", ""))
    sStock = Trim$(Replace(sStock, "Baby ", ""))
    sStock = Trim$(Replace(sStock, "BTL ", ""))
    sStock = Trim$(Replace(sStock, "CAN ", ""))
    sStock = Trim$(Replace(sStock, "CANS ", ""))
    sStock = Trim$(Replace(sStock, "L/N ", ""))
    sStock = Trim$(Replace(sStock, "LN ", ""))
    sStock = Trim$(Replace(sStock, "NAG ", ""))
    sStock = Trim$(Replace(sStock, "PT ", ""))

    If InStr(sStock, " ") <> 0 Then
        sStock = Trim$(Left(sStock, InStr(sStock, " ")))
    End If
    
    SuggestStockItem = sStock
    
End Function


Public Function GetNoOfProductsForPLU(sPLU As String)
Dim rs As Recordset

    On Error GoTo ErrorHandler

    If Val(sPLU) > 0 Then
    
        Set rs = SWdb.OpenRecordset("SELECT COUNT(ID) AS ProdCount FROM tblClientProductPLUs WHERE ClientID = " & Trim$(lSelClientID) & " AND PLUNumber = " & sPLU, dbOpenSnapshot)
        If Not (rs.EOF And rs.BOF) Then
        
            labelNoProds.Caption = rs("ProdCount")
        
        End If
        
        Select Case rs("ProdCount")
            
            Case 1
             labelNoOfProducts = "There is Only        Product linked to this PLU#"
            
            Case Else
             labelNoOfProducts = "    There are           Products linked to this PLU#"
        
        End Select
        
        labelNoOfProducts.Visible = True
        
        rs.Close
    
    End If
    
    GetNoOfProductsForPLU = True

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetNoOfProductsForPLU") Then Resume 0
    Resume CleanExit

End Function

Public Function DeleteProduct(lID As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler

    Set rs = SWdb.OpenRecordset("tblClientProductPLUs")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lID
    If Not rs.NoMatch Then
        rs.Delete
    End If
    ' remove the product
    
    Set rs = SWdb.OpenRecordset("SELECT * FROM tblDeliveries WHERE ClientProdPLUID = " & Trim$(lID))
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do
            rs.Delete
            rs.MoveNext
        Loop While Not rs.EOF
    End If
    ' remove all delivery records for the same product
    
    
    DeleteProduct = True
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("DeleteProduct") Then Resume 0
    Resume CleanExit

End Function

Public Sub SetBoxWidth()
        
    If Len(txtStock) > 4 Then
        txtStock.Width = Len(txtStock) * 185
    Else
        txtStock.Width = txtCost.Width
    End If
    
End Sub

Public Function UpdatePLUNumberOnOtherProducts(lCLId As Long, sOldPLUNo As String, sNewPLUNo As String)

    If lCLId > 0 Then
    ' make sure we've passed some legitimate values
    
        SWdb.Execute "UPDATE tblClientProductPLUs SET PLUNumber = " & sNewPLUNo & " WHERE ClientID = " & Trim$(lCLId) & " AND PLUNumber = " & sOldPLUNo
    
        UpdatePLUNumberOnOtherProducts = True
        
    End If

End Function

Public Sub SetFormSmall(bSmall As Boolean)

    If bSmall Then
        Me.Picture = imgList.ListImages("small").Picture
        Me.Height = 6050
    
    Else
        Me.Picture = imgList.ListImages("big").Picture
        Me.Height = 8910
    
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


Public Function GetQTYSForThisPLU(lCLId As Long, sPLU As String, dblSalesQty As Double, dblSalesQtyDP As Double, dblGlassQty As Double, dblGlassQtyDP As Double)
' ver 530
'Public Function GetQTYSForThisPLU(lCLId As Long, sPLU As String, dblSalesQty As Double, dblSalesQtyDP As Double)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("SELECT * FROM tblClientProductPLUs WHERE ClientID = " & Str$(lCLId) & " AND PLUNumber = " & sPLU & " ORDER BY SalesQty, SalesQtyDP Desc")
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        'rs.MoveFirst

        If Not IsNull(rs("SalesQty")) Then
        ' ver 434 check for isnull
            
            dblSalesQty = rs("SalesQty") + 0
        End If
        
        If Not IsNull(rs("SalesQtyDP")) Then
            dblSalesQtyDP = rs("SalesQtyDP") + 0
        End If
        
'ver530
        If Not IsNull(rs("GlassQty")) Then
            dblGlassQty = rs("GlassQty") + 0
        End If

        If Not IsNull(rs("GlassQtyDP")) Then
            dblGlassQtyDP = rs("GlassQtyDP") + 0
        End If

        
    End If
    rs.Close
    
    GetQTYSForThisPLU = True

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetQTYSForThisPLU") Then Resume 0
    Resume CleanExit

End Function

Public Function PartNotPresentAlready(lProdID As Long, lCLId As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("SELECT * FROM tblClientProductPLUs WHERE ClientID = " & Str$(lCLId) & " AND ProductID = " & Trim$(lProdID))
    If Not (rs.EOF And rs.BOF) Then
        If (lPLUProductID = 0) Or (lPLUProductID <> rs("ID")) Then
        ' so as long as its not an edit
            PartNotPresentAlready = False
        
        Else
            PartNotPresentAlready = True
        
        End If
        
    Else
        PartNotPresentAlready = True
    End If
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("PartNotPresentAlready") Then Resume 0
    Resume CleanExit


End Function

Public Sub SetMeasure(sGls As String)

    If InStr(sGls, "Draught") <> 0 Then
            lblGlass.Visible = True
            txtGlass.Visible = True
            
            lblSell.Caption = "Pint Price"
            lblSellDP.Caption = "Pint Price 2"
    
    ElseIf ((InStr(sGls, "Wine") <> 0) Or (InStr(sGls, "Champ") <> 0)) Then
            lblGlass.Visible = True
            txtGlass.Visible = True
            lblSell.Caption = "Bottle Price"
            lblSellDP.Caption = "Bottle Price 2"
            
    Else
            lblGlass.Visible = False
            txtGlass.Visible = False
            lblGlassDP.Visible = False
            txtGlassDP.Visible = False
    End If
    
'    lblSellDP.Visible = bDualPrice
 '   txtSellDP.Visible = bDualPrice
 '   lblGlassDP.Visible = bDualPrice
 '   txtGlassDP.Visible = bDualPrice
    
    

'        Select Case sGls
'
'        Case 2
'
'            lblGlass.Visible = True
'            txtGlass.Visible = True
'            lblSell.Caption = "Pint Price"
'            lblSellDP.Caption = "Pint Price 2"
'
'        Case 4, 5
'            lblGlass.Visible = True
'            txtGlass.Visible = True
'            lblSell.Caption = "Bottle Price"
'            lblSellDP.Caption = "Bottle Price 2"
'
'        Case Else
'            lblGlass.Visible = False
'            txtGlass.Visible = False
'            lblGlassDP.Visible = False
'            txtGlassDP.Visible = False
'
'        End Select
    
End Sub
