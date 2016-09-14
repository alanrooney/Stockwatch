VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmStockWatch 
   Appearance      =   0  'Flat
   BackColor       =   &H00CDC9CD&
   BorderStyle     =   0  'None
   Caption         =   "Stockwatch"
   ClientHeight    =   13515
   ClientLeft      =   120
   ClientTop       =   -180
   ClientWidth     =   19200
   ControlBox      =   0   'False
   Icon            =   "frmStockWatch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   13515
   ScaleWidth      =   19200
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8LCtl.VSFlexGrid grdClients 
      Height          =   345
      Left            =   1950
      TabIndex        =   8
      ToolTipText     =   "Select Client"
      Top             =   1260
      Visible         =   0   'False
      Width           =   5745
      _cx             =   10134
      _cy             =   609
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
      MousePointer    =   54
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   15459560
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   8421504
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmStockWatch.frx":1CCA
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
      BackColorFrozen =   8421504
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8LCtl.VSFlexGrid grdCount 
      Height          =   4065
      Left            =   18135
      TabIndex        =   131
      Top             =   7140
      Visible         =   0   'False
      Width           =   9195
      _cx             =   16219
      _cy             =   7170
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   15459560
      ForeColor       =   8421504
      BackColorFixed  =   11182762
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483635
      BackColorBkg    =   14669791
      BackColorAlternate=   15200231
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      Begin VB.PictureBox picBars 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   2430
         Picture         =   "frmStockWatch.frx":1D38
         ScaleHeight     =   2295
         ScaleWidth      =   4800
         TabIndex        =   143
         Top             =   585
         Visible         =   0   'False
         Width           =   4800
         Begin VSFlex8LCtl.VSFlexGrid grdBars 
            Height          =   1860
            Left            =   45
            TabIndex        =   144
            Top             =   360
            Width           =   4665
            _cx             =   8229
            _cy             =   3281
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
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmStockWatch.frx":5F62
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
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Bar                   Full    Open        Weight"
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
            Height          =   330
            Left            =   930
            TabIndex        =   145
            Top             =   60
            Width           =   3675
         End
      End
   End
   Begin VB.PictureBox fraPrint 
      BorderStyle     =   0  'None
      Height          =   12345
      Left            =   17460
      Picture         =   "frmStockWatch.frx":5FDA
      ScaleHeight     =   12345
      ScaleWidth      =   13935
      TabIndex        =   123
      Top             =   6240
      Visible         =   0   'False
      Width           =   13935
      Begin VB.PictureBox picSummary 
         BorderStyle     =   0  'None
         Height          =   8985
         Left            =   375
         Picture         =   "frmStockWatch.frx":13866
         ScaleHeight     =   8985
         ScaleWidth      =   9495
         TabIndex        =   132
         Top             =   1440
         Width           =   9495
         Begin VB.TextBox txtNote 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   345
            MultiLine       =   -1  'True
            TabIndex        =   133
            Top             =   5865
            Width           =   8805
         End
         Begin MyCommandButton.MyButton btnSaveNote 
            Height          =   495
            Left            =   6390
            TabIndex        =   134
            Top             =   8265
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   873
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
            BackColorDown   =   3968251
            BackColorOver   =   6805503
            BackColorFocus  =   6805503
            BackColorDisabled=   13748165
            BorderColor     =   32768
            BorderDrawEvent =   1
            BorderWidth     =   0
            TransparentColor=   13486541
            Caption         =   "Save Selected Items && Note"
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
         Begin MyCommandButton.MyButton btnShow 
            Height          =   495
            Left            =   255
            TabIndex        =   140
            Top             =   4380
            Width           =   1635
            _ExtentX        =   2884
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
            TransparentColor=   13486541
            Caption         =   "Show Selected"
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
         Begin VB.Label labelTotal 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5775
            TabIndex        =   139
            Top             =   4470
            Width           =   3330
         End
         Begin VB.Label LabelNote 
            BackStyle       =   0  'Transparent
            Caption         =   "Client Advisory Note"
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
            Height          =   345
            Left            =   270
            TabIndex        =   135
            Top             =   5400
            Width           =   2490
         End
      End
      Begin VB.ComboBox cboShowPrevious 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frmStockWatch.frx":19274
         Left            =   6780
         List            =   "frmStockWatch.frx":19296
         Style           =   2  'Dropdown List
         TabIndex        =   124
         Top             =   30
         Width           =   795
      End
      Begin MyCommandButton.MyButton cmdPrint 
         Height          =   375
         Left            =   9060
         TabIndex        =   127
         Top             =   60
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   661
         BackColor       =   13748165
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
         TransparentColor=   13486541
         Caption         =   "&Print Display"
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
      Begin MyCommandButton.MyButton btnCloseFraPrint 
         Height          =   255
         Left            =   13470
         TabIndex        =   129
         TabStop         =   0   'False
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
         TransparentColor=   13486541
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
      Begin VB.Label lblShowPrevious 
         BackStyle       =   0  'Transparent
         Caption         =   "&Show Previous"
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
         Height          =   345
         Left            =   5160
         TabIndex        =   125
         Top             =   90
         Width           =   2025
      End
      Begin VB.Image imgReport 
         Height          =   9345
         Left            =   1890
         Picture         =   "frmStockWatch.frx":192C4
         Stretch         =   -1  'True
         Top             =   1380
         Visible         =   0   'False
         Width           =   7485
      End
      Begin VB.Label labelTitle 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   285
         TabIndex        =   126
         Top             =   60
         Width           =   3885
      End
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   750
      Top             =   11535
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picCash 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   7635
      Left            =   17445
      Picture         =   "frmStockWatch.frx":21048
      ScaleHeight     =   7635
      ScaleWidth      =   11355
      TabIndex        =   72
      Top             =   4935
      Visible         =   0   'False
      Width           =   11355
      Begin VB.TextBox txtSurpluslbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C5DDFE&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   390
         Left            =   6675
         MaxLength       =   25
         TabIndex        =   99
         Text            =   "Cash €"
         Top             =   2280
         Width           =   2070
      End
      Begin VB.TextBox txtSurplus 
         BackColor       =   &H00C5DDFE&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8745
         TabIndex        =   100
         Text            =   " "
         Top             =   2280
         Width           =   1545
      End
      Begin VB.TextBox txtActual 
         BackColor       =   &H00C5DDFE&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3885
         TabIndex        =   74
         Text            =   " "
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtStaff 
         BackColor       =   &H00C5DDFE&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3885
         TabIndex        =   76
         Text            =   " "
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtComplimentary 
         BackColor       =   &H00C5DDFE&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3885
         TabIndex        =   78
         Text            =   " "
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtWastage 
         BackColor       =   &H00C5DDFE&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3885
         TabIndex        =   80
         Text            =   " "
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtOverRings 
         BackColor       =   &H00C5DDFE&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3885
         TabIndex        =   82
         Text            =   " "
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox txtPromotions 
         BackColor       =   &H00C5DDFE&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3885
         TabIndex        =   84
         Text            =   " "
         Top             =   4050
         Width           =   1335
      End
      Begin VB.TextBox txtOffLicense 
         BackColor       =   &H00C5DDFE&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3885
         TabIndex        =   86
         Text            =   " "
         Top             =   4650
         Width           =   1335
      End
      Begin VB.TextBox txtVoucherSales 
         BackColor       =   &H00C5DDFE&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3885
         TabIndex        =   88
         Text            =   " "
         Top             =   5250
         Width           =   1335
      End
      Begin VB.TextBox txtKitchen 
         BackColor       =   &H00C5DDFE&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3885
         TabIndex        =   90
         Text            =   " "
         Top             =   5820
         Width           =   1335
      End
      Begin VB.TextBox txtOther 
         BackColor       =   &H00C5DDFE&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3885
         TabIndex        =   93
         Text            =   " "
         Top             =   6390
         Width           =   1335
      End
      Begin VB.TextBox txtOtherLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C5DDFE&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   1290
         MaxLength       =   25
         TabIndex        =   92
         Text            =   "Other"
         Top             =   6390
         Width           =   2580
      End
      Begin MyCommandButton.MyButton cmdCashSave 
         Height          =   495
         Left            =   9270
         TabIndex        =   101
         Top             =   6255
         Width           =   930
         _ExtentX        =   1640
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
         TransparentColor=   13486541
         Caption         =   "Sa&ve"
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
      Begin MyCommandButton.MyButton btnCashClose 
         Height          =   255
         Left            =   10980
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   60
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
         TransparentColor=   13486541
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
      Begin VB.Label lblSurplusLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "&Surplus"
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
         Height          =   285
         Left            =   5790
         TabIndex        =   98
         Top             =   2310
         Width           =   930
      End
      Begin VB.Label lblOther 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Other"
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
         Height          =   285
         Left            =   1515
         TabIndex        =   91
         Top             =   6420
         Width           =   2325
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   120
         Top             =   30
         Width           =   645
      End
      Begin VB.Label lblActual 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Actual Cash Takings"
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
         Height          =   375
         Left            =   1575
         TabIndex        =   73
         Top             =   1095
         Width           =   2265
      End
      Begin VB.Label lblStaff 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Staff &Drinks"
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
         Height          =   285
         Left            =   1935
         TabIndex        =   75
         Top             =   1695
         Width           =   1875
      End
      Begin VB.Label lblComplimentary 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Complimentary Drinks"
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
         Height          =   285
         Left            =   1215
         TabIndex        =   77
         Top             =   2310
         Width           =   2625
      End
      Begin VB.Label lblWastage 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Wastage Allowance"
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
         Height          =   345
         Left            =   1290
         TabIndex        =   79
         Top             =   2895
         Width           =   2565
      End
      Begin VB.Label lblOverRings 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Mistakes"
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
         Height          =   345
         Left            =   1965
         TabIndex        =   81
         Top             =   3510
         Width           =   1875
      End
      Begin VB.Label lblPromotions 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Promotions"
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
         Height          =   285
         Left            =   1965
         TabIndex        =   83
         Top             =   4065
         Width           =   1875
      End
      Begin VB.Label lblOffLicense 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Off &License Difference"
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
         Height          =   285
         Left            =   1065
         TabIndex        =   85
         Top             =   4680
         Width           =   2805
      End
      Begin VB.Label lblVoucherSales 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Voucher Sales"
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
         Height          =   285
         Left            =   1515
         TabIndex        =   87
         Top             =   5265
         Width           =   2325
      End
      Begin VB.Label lblCalculatedSales 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Calculated Sales €"
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
         Height          =   285
         Left            =   6465
         TabIndex        =   94
         Top             =   1140
         Width           =   2235
      End
      Begin VB.Label labelCalculatedSales 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   405
         Left            =   8745
         TabIndex        =   95
         Top             =   1080
         Width           =   1545
      End
      Begin VB.Label lblCalculatedActual 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Projected Sales €"
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
         Height          =   285
         Left            =   6615
         TabIndex        =   96
         Top             =   1725
         Width           =   2085
      End
      Begin VB.Label labelProjectedSales 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   405
         Left            =   8745
         TabIndex        =   97
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Label lblKitchen 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Kitchen"
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
         Height          =   285
         Left            =   1515
         TabIndex        =   89
         Top             =   5835
         Width           =   2325
      End
      Begin VB.Label lblOtherLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "&Other"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C5DDFE&
         Height          =   285
         Left            =   1335
         TabIndex        =   119
         Top             =   6450
         Width           =   555
      End
   End
   Begin VB.PictureBox fraSales 
      BorderStyle     =   0  'None
      Height          =   6105
      Left            =   17400
      Picture         =   "frmStockWatch.frx":26D3A
      ScaleHeight     =   6105
      ScaleWidth      =   13350
      TabIndex        =   59
      Top             =   3780
      Visible         =   0   'False
      Width           =   13355
      Begin VB.TextBox txtGlass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2850
         TabIndex        =   65
         Top             =   2175
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtGlassDP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6750
         TabIndex        =   69
         Top             =   2190
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtSalesDP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6750
         TabIndex        =   67
         Top             =   1620
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtSales 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2865
         TabIndex        =   63
         Top             =   1605
         Width           =   765
      End
      Begin MyCommandButton.MyButton cmdTillSave 
         Height          =   495
         Left            =   9690
         TabIndex        =   70
         Top             =   2190
         Width           =   930
         _ExtentX        =   1640
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
         TransparentColor=   13486541
         Caption         =   "Sa&ve"
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
      Begin MyCommandButton.MyButton btnSalesClose 
         Height          =   255
         Left            =   12960
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   60
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
         TransparentColor=   13486541
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
      Begin VB.Label lblGlass 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Glass Qty"
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
         Height          =   285
         Left            =   990
         TabIndex        =   64
         Top             =   2235
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblGlassDP 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "G&lass Qty 2"
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
         Height          =   285
         Left            =   4860
         TabIndex        =   68
         Top             =   2250
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.Label labelGlass1Price 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3795
         TabIndex        =   148
         Top             =   2175
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label labelGlassDPPrice 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7665
         TabIndex        =   147
         Top             =   2190
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label labelSalesDPPrice 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7665
         TabIndex        =   142
         Top             =   1620
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label labelSales1Price 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3795
         TabIndex        =   141
         Top             =   1605
         Width           =   1230
      End
      Begin VB.Label lblSalesDP 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Full Qty 2"
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
         Height          =   285
         Left            =   4860
         TabIndex        =   66
         Top             =   1680
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.Label labelTillCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   390
         Left            =   9345
         TabIndex        =   118
         Top             =   1650
         Width           =   3180
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   117
         Top             =   60
         Width           =   735
      End
      Begin VB.Label lblSales 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Full &Qty"
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
         Height          =   285
         Left            =   975
         TabIndex        =   62
         Top             =   1665
         Width           =   1815
      End
      Begin VB.Label labelTillDescription 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2880
         TabIndex        =   61
         Top             =   1050
         Width           =   7710
      End
      Begin VB.Label lblTillCode 
         BackStyle       =   0  'Transparent
         Caption         =   "Till Item"
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
         Height          =   285
         Left            =   1980
         TabIndex        =   60
         Top             =   1110
         Width           =   1635
      End
   End
   Begin VB.PictureBox fraDelivery 
      BorderStyle     =   0  'None
      Height          =   8655
      Left            =   16815
      Picture         =   "frmStockWatch.frx":2DFC3
      ScaleHeight     =   8655
      ScaleWidth      =   13350
      TabIndex        =   45
      Top             =   2700
      Visible         =   0   'False
      Width           =   13355
      Begin VB.TextBox txtDelOther 
         BackColor       =   &H00FFFFFF&
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
         Height          =   420
         Left            =   7260
         MaxLength       =   50
         TabIndex        =   50
         Top             =   1110
         Visible         =   0   'False
         Width           =   3368
      End
      Begin VB.TextBox txtFree 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3870
         TabIndex        =   56
         Top             =   2190
         Width           =   765
      End
      Begin VB.TextBox txtDelCost 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6300
         TabIndex        =   54
         Top             =   1650
         Width           =   945
      End
      Begin VB.TextBox txtQty 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3870
         TabIndex        =   52
         Top             =   1650
         Width           =   765
      End
      Begin VB.ComboBox cboDelivery 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frmStockWatch.frx":34DA8
         Left            =   3870
         List            =   "frmStockWatch.frx":34DB2
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   1110
         Width           =   3368
      End
      Begin MyCommandButton.MyButton btnDeliveryClose 
         Height          =   255
         Left            =   12960
         TabIndex        =   58
         TabStop         =   0   'False
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
         TransparentColor=   13486541
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
      Begin MyCommandButton.MyButton cmdDelSave 
         Height          =   495
         Left            =   9690
         TabIndex        =   57
         Top             =   2190
         Width           =   930
         _ExtentX        =   1640
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
         TransparentColor=   13486541
         Caption         =   "Sa&ve"
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deliveries"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   116
         Top             =   60
         Width           =   1185
      End
      Begin VB.Label lblFree 
         BackStyle       =   0  'Transparent
         Caption         =   "&Free"
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
         Height          =   285
         Left            =   3360
         TabIndex        =   55
         Top             =   2250
         Width           =   1725
      End
      Begin VB.Label labelDelCount 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   360
         Left            =   4890
         TabIndex        =   115
         Top             =   2220
         Width           =   4590
      End
      Begin VB.Label lblDelCost 
         BackStyle       =   0  'Transparent
         Caption         =   "&Cost"
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
         Height          =   285
         Left            =   5760
         TabIndex        =   53
         Top             =   1710
         Width           =   1725
      End
      Begin VB.Label lblQty 
         BackStyle       =   0  'Transparent
         Caption         =   "&Quantity Received"
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
         Height          =   285
         Left            =   2970
         TabIndex        =   51
         Top             =   1680
         Width           =   1725
      End
      Begin VB.Label lblDelivery 
         BackStyle       =   0  'Transparent
         Caption         =   "&Delivery Note"
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
         Height          =   285
         Left            =   2460
         TabIndex        =   48
         Top             =   1170
         Width           =   2055
      End
      Begin VB.Label lblDelItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Item"
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
         Height          =   285
         Left            =   2460
         TabIndex        =   46
         Top             =   600
         Width           =   1995
      End
      Begin VB.Label labelDelItem 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3870
         TabIndex        =   47
         Top             =   540
         Width           =   6735
      End
   End
   Begin VB.PictureBox fraStock 
      BorderStyle     =   0  'None
      Height          =   12165
      Left            =   4980
      Picture         =   "frmStockWatch.frx":34DCD
      ScaleHeight     =   12165
      ScaleWidth      =   13335
      TabIndex        =   28
      ToolTipText     =   "Stock Details"
      Top             =   1740
      Visible         =   0   'False
      Width           =   13335
      Begin VB.ComboBox cboBar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frmStockWatch.frx":3F244
         Left            =   975
         List            =   "frmStockWatch.frx":3F246
         Style           =   2  'Dropdown List
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   540
         Width           =   2325
      End
      Begin VB.TextBox txtFullQty 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5145
         TabIndex        =   34
         Top             =   1110
         Width           =   765
      End
      Begin VB.TextBox txtOpen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5145
         TabIndex        =   36
         Top             =   1650
         Width           =   765
      End
      Begin VB.TextBox txtWeight 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5145
         TabIndex        =   38
         Top             =   2190
         Width           =   765
      End
      Begin MyCommandButton.MyButton btnClose 
         Height          =   255
         Left            =   12960
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   60
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
         TransparentColor=   13486541
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
      Begin MyCommandButton.MyButton cmdSaveStock 
         Height          =   495
         Left            =   10980
         TabIndex        =   43
         Top             =   2130
         Width           =   930
         _ExtentX        =   1640
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
         TransparentColor=   13486541
         Caption         =   "Sa&ve"
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
      Begin VB.Label lblBar 
         BackStyle       =   0  'Transparent
         Caption         =   "&Bar"
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
         Height          =   285
         Left            =   555
         TabIndex        =   29
         Top             =   585
         Width           =   2055
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Item"
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
         Height          =   285
         Left            =   4380
         TabIndex        =   31
         Top             =   585
         Width           =   705
      End
      Begin VB.Label labelDescription 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5145
         TabIndex        =   32
         Top             =   525
         Width           =   6735
      End
      Begin VB.Label lblFullQty 
         BackStyle       =   0  'Transparent
         Caption         =   "&Full Quantity"
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
         Height          =   285
         Left            =   3795
         TabIndex        =   33
         Top             =   1170
         Width           =   1335
      End
      Begin VB.Label lblOpen 
         BackStyle       =   0  'Transparent
         Caption         =   "&Open Items"
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
         Height          =   285
         Left            =   3855
         TabIndex        =   35
         Top             =   1710
         Width           =   1335
      End
      Begin VB.Label lblWeight 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Weight"
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
         Height          =   345
         Left            =   3060
         TabIndex        =   37
         Top             =   2250
         Width           =   2025
      End
      Begin VB.Label lblFullWeight 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Full Weight"
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
         Height          =   315
         Left            =   5955
         TabIndex        =   39
         Top             =   1170
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.Label labelFullWeight 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   420
         Left            =   8325
         TabIndex        =   40
         Top             =   1110
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lblEmptyWeight 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Empty Weight"
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
         Height          =   315
         Left            =   8625
         TabIndex        =   41
         Top             =   1170
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.Label labelEmptyWeight 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   420
         Left            =   10995
         TabIndex        =   42
         Top             =   1110
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label labelStockCount 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   360
         Left            =   6195
         TabIndex        =   114
         Top             =   2250
         Width           =   4590
      End
      Begin VB.Label labelMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   6015
         TabIndex        =   113
         Top             =   1710
         Width           =   4935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   112
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picMenu 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      Picture         =   "frmStockWatch.frx":3F248
      ScaleHeight     =   435
      ScaleWidth      =   19200
      TabIndex        =   122
      Top             =   0
      Width           =   19200
      Begin MyCommandButton.MyButton btnProducts 
         Height          =   270
         Left            =   2610
         TabIndex        =   2
         Top             =   60
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   476
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         MaskColor       =   65535
         Picture         =   "frmStockWatch.frx":42F38
         BackColorDown   =   13609135
         BackColorOver   =   13609135
         BackColorFocus  =   13609135
         BackColorDisabled=   -2147483633
         BorderColor     =   8323072
         BorderDrawEvent =   6
         BorderWidth     =   0
         TransparentColor=   13486541
         Caption         =   "&Products"
         CaptionPosition =   4
         ForeColorDisabled=   -2147483630
         ForeColorOver   =   13003064
         ForeColorFocus  =   13003064
         ForeColorDown   =   13003064
         PictureDisabled =   "frmStockWatch.frx":43896
         PictureOver     =   "frmStockWatch.frx":4480C
         PictureAlignment=   4
         PictureOverEffect=   2
      End
      Begin MyCommandButton.MyButton btnExit 
         Height          =   270
         Left            =   450
         TabIndex        =   0
         Top             =   60
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   476
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         MaskColor       =   65535
         Picture         =   "frmStockWatch.frx":45782
         BackColorDown   =   13609135
         BackColorOver   =   13609135
         BackColorFocus  =   13609135
         BackColorDisabled=   -2147483633
         BorderColor     =   8323072
         BorderDrawEvent =   6
         BorderWidth     =   0
         TransparentColor=   13486541
         Caption         =   "&Exit"
         CaptionPosition =   4
         ForeColorDisabled=   8421504
         ForeColorOver   =   13003064
         ForeColorFocus  =   13003064
         ForeColorDown   =   13003064
         PictureDisabled =   "frmStockWatch.frx":460E0
         PictureOver     =   "frmStockWatch.frx":47056
         PictureAlignment=   4
         PictureOverEffect=   2
      End
      Begin MyCommandButton.MyButton btnClients 
         Height          =   270
         Left            =   1530
         TabIndex        =   1
         Top             =   60
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   476
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         MaskColor       =   65535
         Picture         =   "frmStockWatch.frx":47FCC
         BackColorDown   =   13609135
         BackColorOver   =   13609135
         BackColorFocus  =   13609135
         BackColorDisabled=   -2147483633
         BorderColor     =   8323072
         BorderDrawEvent =   6
         BorderWidth     =   0
         TransparentColor=   13486541
         Caption         =   "&Clients"
         CaptionPosition =   4
         ForeColorDisabled=   -2147483630
         ForeColorOver   =   13003064
         ForeColorFocus  =   13003064
         ForeColorDown   =   13003064
         PictureDisabled =   "frmStockWatch.frx":4892A
         PictureOver     =   "frmStockWatch.frx":498A0
         PictureAlignment=   4
         PictureOverEffect=   2
      End
      Begin MyCommandButton.MyButton btnPlus 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   3675
         TabIndex        =   3
         Top             =   60
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   476
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         MaskColor       =   65535
         Picture         =   "frmStockWatch.frx":4A816
         BackColorDown   =   13609135
         BackColorOver   =   13609135
         BackColorFocus  =   13609135
         BackColorDisabled=   -2147483633
         BorderColor     =   8323072
         BorderDrawEvent =   6
         BorderWidth     =   0
         TransparentColor=   13486541
         Caption         =   "P&LUs"
         CaptionPosition =   4
         ForeColorDisabled=   -2147483630
         ForeColorOver   =   13003064
         ForeColorFocus  =   13003064
         ForeColorDown   =   13003064
         PictureDisabled =   "frmStockWatch.frx":4B174
         PictureOver     =   "frmStockWatch.frx":4C0EA
         PictureAlignment=   4
         PictureOverEffect=   2
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton btnGroups 
         Height          =   270
         Left            =   4770
         TabIndex        =   4
         Top             =   60
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   476
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         MaskColor       =   65535
         Picture         =   "frmStockWatch.frx":4D060
         BackColorDown   =   13609135
         BackColorOver   =   13609135
         BackColorFocus  =   13609135
         BackColorDisabled=   -2147483633
         BorderColor     =   8323072
         BorderDrawEvent =   6
         BorderWidth     =   0
         TransparentColor=   13486541
         Caption         =   "&Groups"
         CaptionPosition =   4
         ForeColorDisabled=   -2147483630
         ForeColorOver   =   13003064
         ForeColorFocus  =   13003064
         ForeColorDown   =   13003064
         PictureDisabled =   "frmStockWatch.frx":4D9BE
         PictureOver     =   "frmStockWatch.frx":4E934
         PictureAlignment=   4
         PictureOverEffect=   2
      End
      Begin MyCommandButton.MyButton btnSettings 
         Height          =   270
         Left            =   5850
         TabIndex        =   5
         Top             =   60
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   476
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         MaskColor       =   65535
         Picture         =   "frmStockWatch.frx":4F8AA
         BackColorDown   =   13609135
         BackColorOver   =   13609135
         BackColorFocus  =   13609135
         BackColorDisabled=   -2147483633
         BorderColor     =   8323072
         BorderDrawEvent =   6
         BorderWidth     =   0
         TransparentColor=   13486541
         Caption         =   "&Settings"
         CaptionPosition =   4
         ForeColorDisabled=   -2147483630
         ForeColorOver   =   13003064
         ForeColorFocus  =   13003064
         ForeColorDown   =   13003064
         PictureDisabled =   "frmStockWatch.frx":50208
         PictureOver     =   "frmStockWatch.frx":5117E
         PictureAlignment=   4
         PictureOverEffect=   2
      End
      Begin MyCommandButton.MyButton btnAbout 
         Height          =   270
         Left            =   7995
         TabIndex        =   6
         Top             =   60
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   476
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         MaskColor       =   65535
         Picture         =   "frmStockWatch.frx":520F4
         BackColorDown   =   13609135
         BackColorOver   =   13609135
         BackColorFocus  =   13609135
         BackColorDisabled=   -2147483633
         BorderColor     =   8323072
         BorderDrawEvent =   6
         BorderWidth     =   0
         TransparentColor=   13486541
         Caption         =   "A&bout"
         CaptionPosition =   4
         ForeColorDisabled=   -2147483630
         ForeColorOver   =   13003064
         ForeColorFocus  =   13003064
         ForeColorDown   =   13003064
         PictureDisabled =   "frmStockWatch.frx":52A52
         PictureOver     =   "frmStockWatch.frx":539C8
         PictureAlignment=   4
         PictureOverEffect=   2
      End
      Begin MyCommandButton.MyButton btnEnd 
         Height          =   255
         Left            =   18810
         TabIndex        =   130
         Top             =   60
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   450
         BackColor       =   13748165
         ForeColor       =   8421504
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
         TransparentColor=   13486541
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
      Begin MyCommandButton.MyButton btnAudits 
         Height          =   270
         Left            =   6915
         TabIndex        =   136
         Top             =   60
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   476
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         MaskColor       =   65535
         Picture         =   "frmStockWatch.frx":5493E
         BackColorDown   =   13609135
         BackColorOver   =   13609135
         BackColorFocus  =   13609135
         BackColorDisabled=   -2147483633
         BorderColor     =   8323072
         BorderDrawEvent =   6
         BorderWidth     =   0
         TransparentColor=   13486541
         Caption         =   "&Audits"
         CaptionPosition =   4
         ForeColorDisabled=   -2147483630
         ForeColorOver   =   13003064
         ForeColorFocus  =   13003064
         ForeColorDown   =   13003064
         PictureDisabled =   "frmStockWatch.frx":5529C
         PictureOver     =   "frmStockWatch.frx":56212
         PictureAlignment=   4
         PictureOverEffect=   2
      End
      Begin MyCommandButton.MyButton btnMinimize 
         Height          =   255
         Left            =   18450
         TabIndex        =   146
         Top             =   60
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   450
         BackColor       =   13748165
         ForeColor       =   8421504
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
         TransparentColor=   13486541
         Caption         =   "_"
         CaptionOffsetY  =   -5
         CaptionPosition =   4
         ForeColorDisabled=   8421504
         ForeColorOver   =   192
         ForeColorFocus  =   13003064
         ForeColorDown   =   192
         PictureAlignment=   4
         GradientType    =   2
         TextFadeToColor =   8388608
      End
   End
   Begin VB.PictureBox picStatus 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   16000
      Left            =   -15
      Picture         =   "frmStockWatch.frx":57188
      ScaleHeight     =   16005
      ScaleWidth      =   3360
      TabIndex        =   106
      Top             =   1395
      Visible         =   0   'False
      Width           =   3360
      Begin VSFlex8LCtl.VSFlexGrid grdMenu 
         Height          =   7605
         Left            =   180
         TabIndex        =   13
         Top             =   1980
         Width           =   3000
         _cx             =   5292
         _cy             =   13414
         Appearance      =   2
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   54
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   12157534
         ForeColorSel    =   -2147483634
         BackColorBkg    =   8421504
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   8
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   22
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmStockWatch.frx":5D6AF
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
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   120
         Top             =   10140
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   21
         ImageHeight     =   21
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStockWatch.frx":5D880
               Key             =   "button"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStockWatch.frx":5DB76
               Key             =   "tick"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStockWatch.frx":60F62
               Key             =   "cashblank"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStockWatch.frx":6D40C
               Key             =   "stockblank"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStockWatch.frx":7A18F
               Key             =   "salesblank"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStockWatch.frx":889AD
               Key             =   "bkg"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStockWatch.frx":926C4
               Key             =   "inp"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStockWatch.frx":95B98
               Key             =   "deliveryblank"
            EndProperty
         EndProperty
      End
      Begin VB.Label labelDate 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Left            =   1065
         TabIndex        =   138
         Top             =   540
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "On:"
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
         Height          =   345
         Left            =   570
         TabIndex        =   137
         Top             =   540
         Width           =   705
      End
      Begin VB.Label labelTo 
         BackStyle       =   0  'Transparent
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
         Height          =   345
         Left            =   1050
         TabIndex        =   111
         Top             =   1335
         Width           =   1935
      End
      Begin VB.Label labelFrom 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Left            =   1065
         TabIndex        =   110
         Top             =   915
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
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
         Height          =   345
         Left            =   630
         TabIndex        =   109
         Top             =   1350
         Width           =   705
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
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
         Height          =   345
         Left            =   330
         TabIndex        =   108
         Top             =   945
         Width           =   705
      End
      Begin VB.Label labelStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "Count In Progress"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   255
         TabIndex        =   107
         Top             =   195
         Width           =   2835
      End
   End
   Begin VB.PictureBox picSelect 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      Picture         =   "frmStockWatch.frx":A25E2
      ScaleHeight     =   975
      ScaleWidth      =   19140
      TabIndex        =   103
      Top             =   420
      Width           =   19140
      Begin MyCommandButton.MyButton cmdSelectClient 
         Height          =   705
         Left            =   30
         TabIndex        =   7
         ToolTipText     =   "List Clients"
         Top             =   150
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   1244
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
         MaskColor       =   65535
         Picture         =   "frmStockWatch.frx":A669C
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   13748165
         BackColorDisabled=   13748165
         BorderColor     =   32768
         BorderDrawEvent =   1
         BorderWidth     =   0
         TransparentColor=   13486541
         Caption         =   "&Clients"
         CaptionAlignment=   8
         CaptionOffsetX  =   2
         CaptionOffsetY  =   2
         CaptionPosition =   4
         ForeColorDisabled=   -2147483629
         ForeColorOver   =   13003064
         ForeColorFocus  =   13003064
         ForeColorDown   =   13003064
         PictureDisabled =   "frmStockWatch.frx":A70DE
         PictureOffsetX  =   3
         PictureDisabledEffect=   2
         GradientType    =   3
         TextFadeToColor =   8388608
         TextFadeEvents  =   6
      End
      Begin MyCommandButton.MyButton cmdProducts 
         Height          =   705
         Left            =   9450
         TabIndex        =   10
         ToolTipText     =   "Show Client Products"
         Top             =   150
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   1244
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
         MaskColor       =   65535
         Picture         =   "frmStockWatch.frx":A8268
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   13748165
         BackColorDisabled=   13748165
         BorderColor     =   32768
         BorderDrawEvent =   1
         BorderWidth     =   0
         TransparentColor=   13486541
         Caption         =   "&Products "
         CaptionAlignment=   8
         CaptionOffsetY  =   2
         CaptionPosition =   4
         ForeColorDisabled=   8421504
         ForeColorOver   =   13003064
         ForeColorFocus  =   13003064
         ForeColorDown   =   13003064
         PictureDisabled =   "frmStockWatch.frx":A8C36
         GradientType    =   3
         TextFadeToColor =   8388608
         TextFadeEvents  =   6
      End
      Begin MyCommandButton.MyButton cmdReports 
         Height          =   705
         Left            =   10890
         TabIndex        =   11
         ToolTipText     =   "Show Client Reports"
         Top             =   150
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   1244
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
         MaskColor       =   65535
         Picture         =   "frmStockWatch.frx":A9C60
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   13748165
         BackColorDisabled=   13748165
         BorderColor     =   32768
         BorderDrawEvent =   1
         BorderWidth     =   0
         TransparentColor=   13486541
         Caption         =   "&Reports"
         CaptionAlignment=   8
         CaptionOffsetY  =   2
         CaptionPosition =   4
         ForeColorDisabled=   8421504
         ForeColorOver   =   13003064
         ForeColorFocus  =   13003064
         ForeColorDown   =   13003064
         PictureDisabled =   "frmStockWatch.frx":AA5C2
         PictureOffsetX  =   1
         GradientType    =   3
         TextFadeToColor =   8388608
         TextFadeEvents  =   6
      End
      Begin MyCommandButton.MyButton cmdEmail 
         Height          =   705
         Left            =   12315
         TabIndex        =   12
         ToolTipText     =   "Email Reports to Client"
         Top             =   150
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   1244
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
         MaskColor       =   65535
         Picture         =   "frmStockWatch.frx":AB544
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   13748165
         BackColorDisabled=   13748165
         BorderColor     =   32768
         BorderDrawEvent =   1
         BorderWidth     =   0
         TransparentColor=   13486541
         Caption         =   "&Email"
         CaptionAlignment=   8
         CaptionOffsetX  =   4
         CaptionOffsetY  =   2
         CaptionPosition =   4
         ForeColorDisabled=   8421504
         ForeColorOver   =   13003064
         ForeColorFocus  =   13003064
         ForeColorDown   =   13003064
         PictureDisabled =   "frmStockWatch.frx":ABD62
         PictureOffsetX  =   5
         GradientType    =   3
         TextFadeToColor =   8388608
         TextFadeEvents  =   6
      End
      Begin MyCommandButton.MyButton cmdStockTake 
         Height          =   705
         Left            =   8010
         TabIndex        =   9
         ToolTipText     =   "Begin New Audit Count"
         Top             =   150
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   1244
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
         MaskColor       =   65535
         Picture         =   "frmStockWatch.frx":AC918
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   13748165
         BackColorDisabled=   13748165
         BorderColor     =   32768
         BorderWidth     =   0
         TransparentColor=   13486541
         Caption         =   "&Count"
         CaptionAlignment=   8
         CaptionOffsetX  =   3
         CaptionOffsetY  =   2
         CaptionPosition =   4
         DepthMode       =   1
         ForeColorDisabled=   8421504
         ForeColorOver   =   13003064
         ForeColorFocus  =   13003064
         ForeColorDown   =   13003064
         PictureDisabled =   "frmStockWatch.frx":AD2BA
         GradientType    =   3
         TextFadeToColor =   8388608
         TextFadeEvents  =   6
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   525
         Left            =   13950
         TabIndex        =   105
         Top             =   240
         Width           =   4935
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgInProgress 
         Height          =   570
         Left            =   1620
         Picture         =   "frmStockWatch.frx":AE274
         Stretch         =   -1  'True
         ToolTipText     =   "Audit Count In Progress"
         Top             =   210
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblClient 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2025
         TabIndex        =   104
         ToolTipText     =   "Client Name"
         Top             =   180
         Width           =   5775
      End
   End
   Begin MyCommandButton.MyButton MyButton1 
      Height          =   255
      Left            =   18870
      TabIndex        =   121
      Top             =   60
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
      TransparentColor=   13486541
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
   Begin VB.PictureBox fraDates 
      BorderStyle     =   0  'None
      Height          =   5385
      Left            =   16560
      Picture         =   "frmStockWatch.frx":B08B6
      ScaleHeight     =   5385
      ScaleWidth      =   8055
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   8055
      Begin MSComCtl2.MonthView Cal 
         Height          =   2820
         Left            =   2475
         TabIndex        =   25
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
         StartOfWeek     =   16580610
         TitleBackColor  =   11375264
         TitleForeColor  =   16777215
         CurrentDate     =   39972
      End
      Begin VB.CheckBox chkEvaluation 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CEBDC7&
         Caption         =   "Valuation Only"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         TabIndex        =   149
         Top             =   3210
         Width           =   1845
      End
      Begin MSMask.MaskEdBox tedOn 
         Height          =   405
         Left            =   4620
         TabIndex        =   22
         ToolTipText     =   "Beginning Stock Take Date"
         Top             =   2295
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdOnCal 
         Height          =   405
         Left            =   5730
         Picture         =   "frmStockWatch.frx":B5A5F
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Calendar"
         Top             =   2295
         Width           =   405
      End
      Begin VB.CommandButton cmdFromCal 
         Height          =   405
         Left            =   3600
         Picture         =   "frmStockWatch.frx":B5AD3
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Calendar"
         Top             =   1530
         Width           =   405
      End
      Begin VB.CommandButton cmdToCal 
         Height          =   405
         Left            =   5730
         Picture         =   "frmStockWatch.frx":B5B47
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Calendar"
         Top             =   1530
         Width           =   405
      End
      Begin MSMask.MaskEdBox tedFrom 
         Height          =   405
         Left            =   2490
         TabIndex        =   16
         ToolTipText     =   "Beginning Stock Take Date"
         Top             =   1530
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox tedTo 
         Height          =   405
         Left            =   4620
         TabIndex        =   19
         ToolTipText     =   "Ending Stock Take Date"
         Top             =   1530
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MyCommandButton.MyButton cmdCancel 
         Height          =   495
         Left            =   690
         TabIndex        =   26
         ToolTipText     =   "Cancel this Stock Take"
         Top             =   4350
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   873
         BackColor       =   13748165
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
         TransparentColor=   13486541
         Caption         =   "Cancel Stock Take"
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
      Begin MyCommandButton.MyButton cmdStart 
         Height          =   495
         Left            =   5070
         TabIndex        =   24
         ToolTipText     =   "Begin a New Stock for this Client"
         Top             =   4380
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   873
         BackColor       =   13748165
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
         TransparentColor=   13486541
         Caption         =   "&Begin Stock Take"
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
      Begin MyCommandButton.MyButton btnCloseBegin 
         Height          =   255
         Left            =   7680
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   60
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
         TransparentColor=   13486541
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
      Begin VB.Label labelOn 
         BackStyle       =   0  'Transparent
         Caption         =   "On"
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
         Height          =   285
         Left            =   4260
         TabIndex        =   21
         Top             =   2325
         Width           =   1695
      End
      Begin VB.Label lblFrom 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Height          =   285
         Left            =   1860
         TabIndex        =   15
         Top             =   1575
         Width           =   825
      End
      Begin VB.Label lblTo 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Height          =   285
         Left            =   4290
         TabIndex        =   18
         Top             =   1590
         Width           =   825
      End
      Begin VB.Label lblStockTake 
         BackStyle       =   0  'Transparent
         Caption         =   "Begin New Stock Take"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   128
         Top             =   45
         Width           =   3645
      End
   End
End
Attribute VB_Name = "frmStockWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public lPLUID As Long
Public lProductID As Long
Public lPLUGroupID As Long
'Public lPLUProductID As Long

Public lDELProductID As Long
Public lSTKProductID As Long
Public lTillID As Long
Public lStkId As Long
Public lDelID As Long
Public lNextDelID As Long
Public lInvoiceID As Long

'Ver2.2.0
'Public curDefaultCost As Currency

Public bAddIt As Boolean


' Count Variables

Public bFromDate As Boolean
Public bToDate As Boolean
Public bOnDate As Boolean
'Public iMins As Integer

Public bEditOther As Boolean
Public lBarID As Long
Public iGlass As Integer

Private Sub btnAbout_Click()
    
    frmSplash.Show vbModal

    picSelect.Visible = True

End Sub

Private Sub btnAudits_Click()
    
    frmAudits.Show vbModal

    picSelect.Visible = True

    btnClients.Enabled = True

End Sub

Private Sub btnCashClose_Click()
    picCash.Visible = False
    Me.Picture = imgList.ListImages("bkg").Picture

End Sub

Private Sub btnClients_Click()
    
    btnClients.Enabled = False
    
    frmCtrl.bFormIsShown = False
    
    sMenuCtrl = "Clients"
    
    SetupView sMenuCtrl
    
    SetupControl sMenuCtrl
  
    frmCtrl.cboActive.ListIndex = 0
    ' Defaul to Active List

    frmCtrl.Show vbModal

    btnClients.Enabled = True

End Sub

Private Sub btnClose_Click()

    grdCount.Visible = False
    fraStock.Visible = False
    Me.Picture = imgList.ListImages("bkg").Picture

End Sub

Private Sub btnCloseBegin_Click()
    
    fraDates.Visible = False
    Me.Picture = imgList.ListImages("bkg").Picture

End Sub

Private Sub btnClosefraPrint_Click()

    fraPrint.Visible = False
    grdCount.Visible = False
    
End Sub



Private Sub btnDeliveryClose_Click()

    fraDelivery.Visible = False
    grdCount.Visible = False
    Me.Picture = imgList.ListImages("bkg").Picture


End Sub



Private Sub btnEnd_Click()

    End

End Sub

Private Sub btnExit_Click()
    End

End Sub

Private Sub btnGroups_Click()
    
    btnGroups.Enabled = False
    
    frmCtrl.bFormIsShown = False
    
    sMenuCtrl = "Groups"
    
    SetupView sMenuCtrl
    
    SetupControl sMenuCtrl
    
    frmCtrl.optProduct.Value = 1
    
    frmCtrl.txtSearch = ""
    
    frmCtrl.cboActive.ListIndex = 0

    frmCtrl.Show vbModal

    btnGroups.Enabled = True

End Sub

Private Sub btnLater_Click()

End Sub

Private Sub btnNow_Click()

End Sub

Private Sub btnOk_Click()

End Sub

Private Sub btnMinimise_Click()

End Sub

Private Sub btnMinimize_Click()

    Me.WindowState = vbMinimized


End Sub

Private Sub btnPlus_Click()
    
    btnPlus.Enabled = False
    
    frmCtrl.bFormIsShown = False
    
    sMenuCtrl = "PLUs"
    
    SetupView sMenuCtrl
    
    SetupControl sMenuCtrl
    
    frmCtrl.txtSearch = ""

    frmCtrl.cboActive.ListIndex = 0
    ' Defaul to Active List

    frmCtrl.Show vbModal
    
    btnPlus.Enabled = True
    
End Sub

Private Sub btnProducts_Click()
    
    btnProducts.Enabled = False
    
    frmCtrl.bFormIsShown = False
    
    sMenuCtrl = "Products"
    
    SetupView sMenuCtrl
    
    SetupControl sMenuCtrl
    
    frmCtrl.txtSearch = ""
    
    frmCtrl.cboActive.ListIndex = 0
    ' Defaul to Active List

    frmCtrl.Show vbModal

    btnProducts.Enabled = True

End Sub

Private Sub btnSalesClose_Click()
    fraSales.Visible = False
    grdCount.Visible = False
    Me.Picture = imgList.ListImages("bkg").Picture

End Sub

Private Sub btnSaveNote_Click()

'ver 310
' Save Note and products ticked

    gbOk = SaveNoteAndProductsTicked(lDatesID)

End Sub

Private Sub btnSettings_Click()
    
'    picSelect.Visible = False
    
    frmSettings.Show vbModal

    picSelect.Visible = True

End Sub

Private Sub btnShow_Click()
Dim iRow As Integer

    If btnShow.Caption = "Show Selected" Then
    
        ShowSelected
        
        
        btnShow.Caption = "Show All"
    Else
        
        For iRow = 2 To grdCount.Rows - 1
            grdCount.RowHidden(iRow) = False
        Next
        
        btnShow.Caption = "Show Selected"
    
    End If

End Sub

Private Sub Cal_DateClick(ByVal DateClicked As Date)
    
    If bFromDate Then
         tedFrom.Text = Format(Cal.Value, sDMY)
    ElseIf bToDate Then
         tedTo.Text = Format(Cal.Value, sDMY)
    ElseIf bOnDate Then
        tedOn.Text = Format(Cal.Value, sDMY)
    End If
    
    Cal.Visible = False


End Sub

Private Sub cboBar_Click()

    txtFullQty.Text = ""
    txtOpen.Text = ""
    txtWeight.Text = ""
    lblFullQty.Tag = ""
    grdCount.Row = 0
    
    ' init input boxes and label tag
    
' Ver440
' Check to see if a bar is clicked....
    
    If cboBar.ListIndex > 0 Then
        
        grdBars.Visible = False
        
        InitEnableInput True
        ' Since a bar other than 'ALL' selected then we enable
        ' inputting boxes
    
        lBarID = cboBar.ItemData(cboBar.ListIndex)
        ' store the Bar ID
        
        gbOk = ShowBarStockCount(lBarID)
    
        If GetNextStockItem(lStkId) Then
        ' Get next blank stock rec to be filled in
             
            bSetFocus Me, "txtFullQty"
        
        End If
        
    Else
    ' 'All' selected so don't allow figures to be entered.
    ' Clear the bar Id to be on the safe side
    
        grdBars.Visible = True

        
        InitEnableInput False
    
        lBarID = 0
        
        gbOk = ShowTotalStockCount()
        ' show stock to be counted for this client
    
        bSetFocus Me, "grdCount"
        
    
    End If



End Sub

Private Sub cboBar_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub cboBar_LostFocus()
    lblBar.ForeColor = sBlack

End Sub

Private Sub cboDelivery_Click()

    If cboDelivery.ListIndex = 1 Then
        
        SetFreeLocked True
        
        txtDelOther.Visible = True
        bSetFocus Me, "txtDelOther"
    Else
    ' Stock Purchase
        
        txtQty.Text = Replace(txtQty, "-", "")
        ' Make sure no Neg appearing in Quantity for a stock purchase
        
        txtDelOther.Visible = False
    
        SetFreeLocked False
    
    End If

End Sub

Private Sub cboDelivery_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub cboDelivery_LostFocus()
    lblDelivery.ForeColor = sBlack

End Sub



Private Sub cboShowPrevious_Click()
    gbOk = RepTillReconciliation(cboShowPrevious)

End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkEvaluation_Click()

    Select Case chkEvaluation.Value

        Case 1
         lblFrom.Visible = False
         tedFrom.Visible = False
         cmdFromCal.Visible = False
        
         lblTo.Visible = False
         tedTo.Visible = False
         cmdToCal.Visible = False
        
         tedOn.Text = Format(Now, sDMY)
         tedFrom.Text = Format(Now, sDMY)
         tedTo.Text = Format(Now, sDMY)
        
        Case Else
         lblFrom.Visible = True
         tedFrom.Visible = True
         cmdFromCal.Visible = True
        
         lblTo.Visible = True
         tedTo.Visible = True
         cmdToCal.Visible = True
        
    End Select

End Sub

Private Sub cmdCancel_Click()

    If MsgBox("Are you sure you want to Cancel this Stock Take", vbDefaultButton2 + vbYesNo + vbQuestion, "Cancel Stock Take") = vbYes Then
    
        If RemoveClientDates(lDatesID) Then
            lDatesID = 0
            
            If RestoreLastCountFigures(lSelClientID) Then
        
                If SetCountInProgress(lSelClientID, False) Then
        
                    If GetClientDate(lSelClientID, lDatesID) Then
                    ' dates of audit
                    
                        If SetSymbolMsgMenu(lSelClientID) Then

                            If SetBeginEndStockTakeButtons(lSelClientID) Then

                                ShowMenu vbTrue
                                LogMsg Me, "Stock Take Cancelled for " & Replace(lblClient.Tag, "_", " "), ""
        
        

                            End If
                        End If
                        
                    End If
                Else
                    LogMsg Me, "Warning - Could not clear Count In Progress flag", "ClientID:" & Trim$(lSelClientID)
                End If
            Else
                LogMsg Me, "Warning - Could not restore last count figures", "ClientID:" & Trim$(lSelClientID)
            End If
        Else
            LogMsg Me, "Warning - Could not remove Stock Take Dates", "ClientID:" & Trim$(lSelClientID)
        End If
    End If
    

End Sub

Private Sub cmdCashSave_Click()

  If ConfirmClientName() Then
    
    cmdCashSave.Enabled = False
    
    gbOk = SaveCashTakings(lDatesID)
  
    LogMsg Me, "Cash Takings Entered for " & Replace(lblClient.Tag, "_", " "), " From: " & labelFrom & " To: " & labelTo
    
    cmdCashSave.Enabled = True
    
    bSetFocus Me, "grdMenu"
  
  End If
  
End Sub

Private Sub cmdDelSave_Click()
  
    gbOk = DoSave(cboDelivery.ListIndex, Val(grdCount.Cell(flexcpTextDisplay, grdCount.Row, grdCount.ColIndex("ref"))))
  
End Sub

Private Sub cmdEmail_Click()
    
    frmEmail.Show vbModal

End Sub

Private Sub cmdFromCal_Click()
    If Cal.Visible = True Then
        Cal.Visible = False
    Else
        If IsDate(tedFrom) Then
            Cal.Value = tedFrom
        Else
            Cal.Value = Format(Now, sDMY)
        End If
        bFromDate = True
        bToDate = False
        bOnDate = False
        Cal.Top = tedFrom.Top + tedFrom.Height
        Cal.Left = tedFrom.Left
        Cal.Visible = True
    End If


End Sub


Private Sub cmdOnCal_Click()
    If Cal.Visible = True Then
        Cal.Visible = False
    Else
        If IsDate(tedOn) Then
            Cal.Value = tedOn
        Else
            Cal.Value = Format(Now + 1, sDMY)
        End If
        bFromDate = False
        bToDate = False
        bOnDate = True
        Cal.Top = tedOn.Top + tedOn.Height
        Cal.Left = tedOn.Left
        Cal.Visible = True
    End If

End Sub

Private Sub cmdPrint_Click()

    Select Case grdMenu.RowData(grdMenu.Row)
     
        Case "A"
         gbOk = PrintDisplay(frmStockWatch, "grdCount", False)
    
        Case Else
         gbOk = PrintDisplay(frmStockWatch, "grdCount", True)

    End Select
    
End Sub

Private Sub cmdProducts_Click()
    
    sMenuCtrl = "Product/PLUs"
    
    SetupView sMenuCtrl
    
    SetupControl sMenuCtrl
    
    frmCtrl.bClientProducts = True
    
    frmCtrl.grdList.ScrollBars = flexScrollBarVertical
    
    frmCtrl.cboActive.ListIndex = 0
    
'    Pause 2000
    
    frmCtrl.Show
    
End Sub

Private Sub cmdReports_Click()

    frmViewReports.Show vbModal

End Sub

Private Sub cmdSaveStock_Click()
Dim lProdID As Long

  If ConfirmClientName() Then
    
    If StockFieldsOk() Then
        
        cmdSaveStock.Enabled = False
        ' disable this button until done
        
        bHourGlass True
        
        If SaveBarCount(lStkId, lSelClientID, lBarID) Then
        ' Ver 440
        ' This was aded to save the quantities for multiple bars
                
            If SaveStockCount(lStkId) Then
                
                If UpdateStockCount(lStkId) Then
                    
                    LogMsg Me, "Stock Item: " & labelDescription & " Entered for " & Replace(lblClient.Tag, "_", " "), labelDescription & "Full Qty:" & txtFullQty & " Open:" & txtOpen & " Weight:" & txtWeight
                    
                    
                    gbOk = InitStockEntry()
                    
                    labelStockCount.Caption = SetCount(grdCount, "FullQty")
                    
                    If GetNextStockItem(lStkId) Then
        
                        bSetFocus Me, "txtFullQty"
                        
                    Else
                        bSetFocus Me, "grdMenu"
                    End If
                    
                End If
            End If
        End If
    
        cmdSaveStock.Enabled = True
        '
    
    
    End If
  End If
  
  bHourGlass False
  
End Sub

Private Sub cmdSelectClient_Click()
         
    If Not grdClients.Visible Then
    ' toggle Clients list
    
       grdClients.Visible = True
       gbOk = ShowActiveClientList()
       
       grdClients.Row = grdClients.FindRow(lSelClientID)
       
       
       bSetFocus Me, "grdClients"
'       If grdClients.Rows > 0 Then grdClients.Row = 0
       ' Show active list of clients and set focus to it.
    
    Else
    ' toggle client list
       grdClients.Visible = False
    End If
    bHourGlass False

End Sub

Private Sub cmdStockTake_Click()
         
    cmdStockTake.Enabled = False
    
    SetUpStockTakeForms  ' Top, left, width, height
    
'Ver 1.4

    If GetClientDate(lSelClientID, lDatesID) Then
    ' dates of audit
        
'        lInvoiceID = GetInvoiceID(lDatesID)
        lInvoiceID = lDatesID
        ' GET INVOICE ID IF THERE IS ONE
        
        If SetSymbolMsgMenu(lSelClientID) Then

            ShowMenu vbTrue
        
        End If
        
    End If
    
    gbOk = SetBeginEndStockTakeButtons(lSelClientID)
    
    bEvaluation = CheckForEvaluation(lSelClientID)
    
    grdMenu.Row = 1
    bSetFocus Me, "grdMenu"

End Sub

Private Sub cmdTillSave_Click()

  If ConfirmClientName() Then
    
    If TillFieldsOK() Then
        
        
        cmdTillSave.Enabled = False
        
'Ver433 ================================================================
' This function replaces the one commented out below it.
' less likely to cause problems
' problems showed up during dual price client


        If UpdateQtyAllItemsSamePLU(lSelClientID, labelTillDescription.Tag, txtSales, txtSalesDP, txtGlass, txtGlassDP) Then


'        If UpdateQtyOtherItemsSamePLU(lSelClientID, labelTillDescription.Tag, txtSales, txtSalesDP) Then
        ' in case there are other items linked to the same PLU we
        ' must make sure the same Sales Qty is applied to all these items.
        ' (.tag holds te plu number)
        
        
            gbOk = RefreshTillSales(lTillID, lSelClientID, labelTillDescription.Tag)
                
            LogMsg Me, "Till Sales Item: " & labelTillDescription & " Entered for " & Replace(lblClient.Tag, "_", " "), labelTillDescription.Caption & " Qty:" & txtSales
            
            labelTillCount.Caption = SetCount(grdCount, "No")
                
            If GetNextTillNo(lTillID) Then
                bSetFocus Me, "txtSales"
            Else
                bSetFocus Me, "grdMenu"
            
            End If
            
        End If
            
        cmdTillSave.Enabled = True
    
    End If
  End If
  
End Sub

Private Sub cmdToCal_Click()
    If Cal.Visible = True Then
        Cal.Visible = False
    Else
        If IsDate(tedTo) Then
            Cal.Value = tedTo
        Else
            Cal.Value = Format(Now, sDMY)
        End If
        bFromDate = False
        bToDate = True
        bOnDate = False
        Cal.Top = tedTo.Top + tedTo.Height
        Cal.Left = tedTo.Left
        Cal.Visible = True
    End If
    

End Sub

Private Sub cmdStart_Click()

  If chkEvaluation Then

    'ver530
    ' Check Evaluation flag
        
    ' EVALUATION:
    
      
      If OnDateCheck(tedOn) Then
  
        If CountInProgress(lSelClientID) Then
            MsgBox "Count already in progress"
        Else
                
            cmdStart.Enabled = False
                
            If StartCount(0, tedFrom, tedTo, tedOn) Then
            ' save dates to tblDates table
            
                gbOk = SetCountInProgress(lSelClientID, True)
  
                gbOk = SetEvaluationFlag(lSelClientID, True)
                
'                gbOk = SaveTillDifference(lSelClientID, lDatesID)
                    
                gbOk = SaveLastStockFigures(lSelClientID)
                ' Save These numbers before we zero them below
    
                gbOk = ClearClientLastCountFigures(lSelClientID)
                    
                gbOk = ClearClientLastBarFigures(lSelClientID)
  
                gbOk = GetClientDate(lSelClientID, lDatesID)
                ' dates of audit
    
                gbOk = SetSymbolMsgMenu(lSelClientID)
                ' Audit symbol
                ' In Progress message in red
                ' modified Menu
  
                gbOk = SetBeginEndStockTakeButtons(lSelClientID)
                
                SetMenuEvaluation True
            
                bEvaluation = True
            
                cmdStart.Enabled = True
            
                fraDates.Visible = False
                bSetFocus Me, "grdMenu"
            
            
            End If
            
        End If
      End If
    
  Else
    If DatesCheck(tedFrom, tedTo) Then
    
      If OnDateCheck(tedOn) Then
      
        If CountInProgress(lSelClientID) Then
        
            If MsgBox("Are you sure you want to change the dates of this Stock Take?", vbDefaultButton2 + vbYesNo + vbQuestion, "Edit Dates") = vbYes Then
            
                If StartCount(lDatesID, tedFrom, tedTo, tedOn) Then
                ' save dates to tblDates table
            
                    gbOk = GetClientDate(lSelClientID, lDatesID)
                    ' dates of audit
                    
                    LogMsg Me, "Stock Take Dates changed for " & Replace(lblClient.Tag, "_", " ") & " From: " & tedFrom & " To: " & tedTo, ""
                
                End If
            End If
        
        ElseIf MsgBox("Begin a New Stock Take for this Client?", vbDefaultButton1 + vbYesNo + vbQuestion, "Begin Stock Take") = vbYes Then
        ' new stock take ...
        
            If DatesCheck(tedFrom, tedTo) Then
            ' verify dates
            ' warn if gaps etc...
            
            
                cmdStart.Enabled = False
                
                If StartCount(0, tedFrom, tedTo, tedOn) Then
                ' save dates to tblDates table
            
                    gbOk = SetCountInProgress(lSelClientID, True)
                    
                    gbOk = SaveTillDifference(lSelClientID, lDatesID)
                    
                    gbOk = SaveLastStockFigures(lSelClientID)
                    ' Save These numbers before we zero them below
    
                    gbOk = ClearClientLastCountFigures(lSelClientID)
                    
                    gbOk = ClearClientLastBarFigures(lSelClientID)
                    'Ver 440 change to included Bars
                    
                    If CountInProgress(lSelClientID) Then
    
                        If GetClientDate(lSelClientID, lDatesID) Then
                        ' dates of audit
    
                            If SetSymbolMsgMenu(lSelClientID) Then
                            ' Audit symbol
                            ' In Progress message in red
                            ' modified Menu
    
                                gbOk = SetBeginEndStockTakeButtons(lSelClientID)
    
                                ShowMenu vbTrue
    
                                LogMsg Me, "Begin New Stock Take for " & Replace(lblClient.Tag, "_", " ") & " On: " & tedOn & " From: " & tedFrom & " To: " & tedTo, ""
    
                            End If
                        End If
                    End If
                End If
            
                cmdStart.Enabled = True
            
            End If
        End If
        
        fraDates.Visible = False
        bSetFocus Me, "grdMenu"
              
      Else
        LogMsg Me, "Invalid On Date: " & tedOn, ""
      End If
    
    End If
  
  End If
  
End Sub




Private Sub Form_Activate()
    
    SetFormSize
    
    SetuplblMsgWidth
    
                'GET AutoUpdate flag
                
                
    bSetFocus Me, "grdMenu"

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    
    If KeyAscii = 13 Then
    ' return pressed force focus to next available object in tabbing order
        gbOk = GotoNextControl(Me, 0)
    
    ElseIf KeyAscii = 27 Then
    ' <ESC>
        
        KeyAscii = 0

        ' CLIENT LIST
        If grdClients.Visible Then
            grdClients.Visible = False
        
        ' CALENDAR
        ElseIf Cal.Visible Then
            Cal.Visible = False
        
        ' DATES PANEL
        ElseIf fraDates.Visible Then
            fraDates.Visible = False
            grdCount.Visible = False
            Me.Picture = imgList.ListImages("bkg").Picture
        
        ElseIf fraDelivery.Visible Then
            fraDelivery.Visible = False
            grdCount.Visible = False
            Me.Picture = imgList.ListImages("bkg").Picture
        
        ElseIf fraStock.Visible Then
            fraStock.Visible = False
            grdCount.Visible = False
            Me.Picture = imgList.ListImages("bkg").Picture
        
        ElseIf fraSales.Visible Then
            fraSales.Visible = False
            grdCount.Visible = False
            Me.Picture = imgList.ListImages("bkg").Picture
        
        ElseIf picCash.Visible Then
            picCash.Visible = False
            Me.Picture = imgList.ListImages("bkg").Picture
        
        
        ' REPORTS
        ElseIf fraPrint.Visible Then
            
            fraPrint.Visible = False
            grdCount.Visible = False
            Me.Picture = imgList.ListImages("bkg").Picture
            
            
        ' SETTINGS
        ElseIf frmCtrl.Visible Then
            picSelect.Visible = True
            
        ElseIf picStatus.Visible Then
        
            picStatus.Visible = False
            Me.Picture = imgList.ListImages("bkg").Picture

        Else
            InitStockWatch
            bSetFocus Me, "cmdSelectClient"

        End If
    End If

End Sub

Private Sub Form_Load()
Dim sIgnore As String

    ' THIS COPY (5.2.0) IS A COPY OF VER 454 (FROM FOLDER 4.4) WHICH IS WHAT ALL FRANCHISES
    ' ARE CURRENTLY USING.
    
    ' iNTO THIS WILL BE INTEGRATED THE CODE TO LINK TO THE TABLET (FROM VER 5.0.0)
    
    ' SEE STOCKWATCH UPDATE.TXT TO SEE ALL CHANGES IMPLIMENTED IN THIS VERSION.
    
    


    If Not App.PrevInstance Then
    ' as long as its not running already...
        
        bHourGlass True
        ' busy
 
        gbOk = OpenAuditFile(Now)
        ' open audit file early
            
        sDBLoc = "" & GetSetting(App.Title, "DB", App.Title & "DB") & ""
        ' get the DB Location from the registry
        
        '------------------------------------------------------------------
        ' SET UP ENVIRONMENT
        If CurDir$ = "C:\Program Files\Microsoft Visual Studio\VB98" Then
            ' TEST
            gbTestEnvironment = True
        Else
            gbTestEnvironment = False
        End If
        

        '-----------------------------------------------------------------
'        If Not gbTestEnvironment Then gbOk = CheckForNewAgentProgram()
        ' CHECK FOR AN UPDATE
        ' IF StockWatchNEW.exe found - check its version and compare with this prog
        
        ' This check is removed for the moment
        
        If gbOpenDB(Me) Then
        ' now open db
        
            On Error Resume Next
            
            If Not gbTestEnvironment Then
                
                gbOk = RestartAgentProgram()
                ' CHECK that SWIAgent is running
                ' If not start it
            
            End If
            
            ' ver 440
            gbOk = CheckAndApplyAnyUpdate()
            
            gbOk = getRegionEmailAndXferLocation(gbRegion, sIgnore, SW1)
            ' Region needed here before we check for license
            
            ShowSplash 1000
            ' Splash Screen & wait
            
            If NoLicense Then
                
                End
            
            Else
                
                sngvatrate = GetVatRate("S")    ' get standard vat rate
    
                SetUpMenuItemsColoursButtons   ' Setup rowdata for each row in menu
                
                GetEmailDefaults    ' from registry
                
                If GetCountsInProgress(lSelClientID, lDatesID) = 1 Then
    
                    If ShowClient(lSelClientID) Then
                    ' the Client name
    
                        SetCtrlButtons vbTrue
                        ' Enable buttons
    
                        cmdStockTake_Click
                        ' Click the Stock Take Button
                    End If
                End If
                
                Unload frmSplash
                ' we're done ... Unload Splash Screen
                
    
'''                '--------------------------------------------
'''                iMins = 30
'''                ' In 1 minute after start up force sending of
'''                ' audit summary files and invoice emails
'''                ' (there after its every 30 minutes)
'''
'''                timXfer.Enabled = True
'''                '--------------------------------------------
                
                bHourGlass False
            
            End If
        
        End If
        
    Else
        MsgBox App.Title & " already running"
        End
    
    End If


End Sub

Private Sub Form_Resize()

'    frmCtrl.grdList.Width = frmStockWatch.Width - 120
    
    SetuplblMsgWidth

End Sub

Public Sub SetuplblMsgWidth()
    
    picSelect.Width = picMenu.Width
    btnEnd.Left = picSelect.Width - 350
    btnMinimize.Left = picSelect.Width - 710
 '   btnUpdateAvailable.Left = picSelect.Width - 1750
    
    If picMenu.Width > (cmdEmail.Left + cmdEmail.Width + 280) Then
        lblMsg.Width = picMenu.Width - (cmdEmail.Left + cmdEmail.Width + 280)
    End If
    
End Sub





'''Private Sub grdClients_Click()
'''
'''    If grdClients.Row > -1 Then
'''
'''        ShowMenu vbFalse
'''
'''        lSelClientID = grdClients.RowData(grdClients.Row)
'''        ' New Client Selected - Save the ID
'''
'''        lDatesID = 0
'''        ' ver 304 - while adding franchise changes i noticed this id wasnt cleared
'''        ' switching between an In progress Client and one that was finished.
'''        ' To be on the safe side i cleared it
'''
'''        If ShowClient(lSelClientID) Then
'''        ' the Client name
'''
'''    'VER1.4
'''
'''            gbOk = SetSymbolMsgMenu(lSelClientID)
'''
'''            SetCtrlButtons vbTrue
'''            ' Enable buttons
'''
'''            If CountInProgress(lSelClientID) Then
'''
'''                cmdStockTake_Click
'''
'''
'''                bSetFocus Me, "grdMenu"
'''            Else
'''
'''                bSetFocus Me, "cmdStockTake"
'''
'''
'''            End If
'''
'''    '        SetCtrlButtons vbTrue
'''            ' Enable buttons
'''
'''        End If
'''        bHourGlass False
'''
'''        HideAllFrames
'''    End If
'''
'''End Sub

Private Sub grdClients_Click()
Dim sDrv As String
'Dim sNewProducts As String

    If grdClients.Row > -1 Then
    
        ShowMenu vbFalse
    
        lSelClientID = grdClients.RowData(grdClients.Row)
        ' New Client Selected - Save the ID
        
        lDatesID = 0
        ' ver 304 - while adding franchise changes i noticed this id wasnt cleared
        ' switching between an In progress Client and one that was finished.
        ' To be on the safe side i cleared it
        
        If ShowClient(lSelClientID) Then
        ' the Client name
        
    'VER1.4
            
            gbOk = SetSymbolMsgMenu(lSelClientID)
            
            SetCtrlButtons vbTrue
            ' Enable buttons
            
            If CountInProgress(lSelClientID) Then
                
                
                ' ver 500
                ' If import file found for this client then
                ' show message and change item 4 in menu to
                ' 'Import Tablet Count'
                
                If TabletImportFileFound(lSelClientID, sDrv) Then
                
                    If MsgBox("Tablet Stock Count File Found for this Client. Import Now?", vbDefaultButton1 + vbYesNo + vbQuestion, "Tablet Import File Found") = vbYes Then
                    
                        
                        
                        ' HERE CHECK FOR NEW PRODUCTS FILE FOR THIS CLIENT
                                                    
                        If ImportNewProductsFound(lSelClientID, sDrv) Then
                        
                            ' IF PRESENT SHOW ARRAY LISTING PRODUCTS
                            
                            frmNewSWCountItems.sNewFileName = sDrv & "SWCount_" & Trim$(lSelClientID) & "NEW_PRODUCTS.csv"
                            frmNewSWCountItems.sDrive = sDrv
                            frmNewSWCountItems.Show
                            
                        Else
                        
                            If ImportTabletFile(lSelClientID, sDrv) Then
                        
                                If DeleteImportedFiles(lSelClientID, sDrv) Then
                                
                                    MsgBox "Import Complete"
                                End If

                            End If
                        
                        End If
                        
                        
                        
                        
                        
                        
                    End If
                    
                Else
                    cmdStockTake_Click
                End If
                
                bSetFocus Me, "grdMenu"
            Else
        
                bSetFocus Me, "cmdStockTake"
            
            
            End If
        
    '        SetCtrlButtons vbTrue
            ' Enable buttons
        
        End If
        bHourGlass False
        
        HideAllFrames
    End If
    
End Sub


Private Sub grdClients_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then grdClients_Click


End Sub

Private Sub grdClients_LostFocus()

    grdClients.Visible = False

End Sub

Private Sub grdCount_Click()

    If grdCount.Row > grdCount.FixedRows - 1 Then
    
        If grdCount.RowData(grdCount.Row) <> Empty Or grdCount.Cell(flexcpData, grdCount.Row, 0) <> Empty Then
        
            Select Case grdMenu.RowData(grdMenu.Row)
        
                Case 4  ' Stock Item
                 
                 lStkId = grdCount.RowData(grdCount.Row)
                 gbOk = GetStockItem(lStkId)
                 bSetFocus Me, "txtFullQty"
'ver440 out for now
'                 If grdCount.ColKey(grdCount.Col) = "Chk" Then
'                    If grdCount.Cell(flexcpData, grdCount.Row, grdCount.Col) = False Then
'                        grdCount.Cell(flexcpData, grdCount.Row, grdCount.Col) = True
'                        grdCount.Cell(flexcpPicture, grdCount.Row, grdCount.Col) = imgList.ListImages("tick").Picture
'                        grdCount.Cell(flexcpPictureAlignment, grdCount.Row, grdCount.Col) = 4
'                    Else
'                        grdCount.Cell(flexcpData, grdCount.Row, grdCount.Col) = False
'                        grdCount.Cell(flexcpPicture, grdCount.Row, grdCount.Col) = ""
'
'                    End If
'                    gbOk = SaveTick(lStkId, cboBar.ItemData(cboBar.ListIndex), grdCount.Cell(flexcpData, grdCount.Row, grdCount.Col))
'                 End If
        
                Case 5  ' Stock Deliveries
                 
                 gbOk = InitDeliveryItem()
                 
                 If Not IsNull(grdCount.Cell(flexcpData, grdCount.Row, 0)) Then
                    lDelID = grdCount.Cell(flexcpData, grdCount.Row, 0)
                 Else
                    lDelID = 0
                 End If
                 
                 lNextDelID = grdCount.RowData(grdCount.Row)
                 
                 If lDelID <> 0 Then
                 ' see if its an edit of a delivered item
                    
                    gbOk = GetDeliveryItem(True, lDelID)
                 
                    If grdCount.Cell(flexcpTextDisplay, grdCount.Row, 0) = 2 Then
                        cmdDelSave.Caption = "&Delete"
                        cmdDelSave.Enabled = True
                        SetDeliveryLocked True
                    
                    Else
                        SetDeliveryLocked False
                    End If
                    
                    bEditOther = False
                    
                    bSetFocus Me, "txtQty"
    
                 Else
                 ' new delivery ...
                    
                    gbOk = GetDeliveryItem(False, lNextDelID)
                 
                    bSetFocus Me, "txtQty"
                 
                 End If
                 
        
                Case 6  ' Till Sales
                 
                 lTillID = grdCount.RowData(grdCount.Row)
                 gbOk = GetPLUNo(lTillID)
                 bSetFocus Me, "txtSales"
        
                Case "G"
                 If grdCount.Cell(flexcpChecked, grdCount.Row, grdCount.ColIndex("Sel")) = flexChecked Then
                    grdCount.Cell(flexcpChecked, grdCount.Row, grdCount.ColIndex("Sel")) = flexUnchecked
                 Else
                    grdCount.Cell(flexcpChecked, grdCount.Row, grdCount.ColIndex("Sel")) = flexChecked
                 End If
                 
' ver 542 removed
'                  ElseIf SelectedMax() Then
'                    grdCount.Cell(flexcpChecked, grdCount.Row, grdCount.ColIndex("Sel")) = flexChecked
'                 Else
'                    MsgBox "Cannot Select more than 25"
'                    grdCount.Cell(flexcpChecked, grdCount.Row, grdCount.ColIndex("Sel")) = flexUnchecked
'                 End If
            
                 gbOk = totalSelected()
                 
                 If btnShow.Caption = "Show All" Then ShowSelected
                 
                 SaveSummaryAndNote True
                 

            End Select
        
        End If
    End If
    
End Sub

Private Sub grdCount_GotFocus()

    If grdCount.Rows > 1 Then
                
        If grdCount.Row = 0 Then grdCount.Row = 1
    End If
    

End Sub


Private Sub grdCount_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        grdCount_Click
    
    End If
    
End Sub

Private Sub grdCount_KeyUp(KeyCode As Integer, Shift As Integer)

    If grdCount.Row > -1 Then
    
        If grdCount.RowData(grdCount.Row) <> Empty Or grdCount.Cell(flexcpData, grdCount.Row, 0) <> Empty Then
        
            Select Case grdMenu.RowData(grdMenu.Row)
        
                Case 4  ' Stock Item
                 gbOk = InitStockEntry()
                 ' clear the entry screen
        
                Case 5  ' Stock Deliveries
                 
                 gbOk = InitDeliveryItem()
                 
                Case 6  ' Till Sales
                 
                 gbOk = InitTillItem()
            
            End Select
        
        End If
    
    End If
    
End Sub

Private Sub grdCount_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Ver 440 New subroutine added here
    ' This basically shows a popup panel with the counts for each bar of the particular
    ' item clicked. It only works while 'All Bars' is being viewed.
    
    picBars.Visible = False

    If grdMenu.RowData(grdMenu.Row) = 4 Then
    ' stock count screen
    
      If grdCount.RowData(grdCount.Row) <> "" Then
    
        If bMultipleBars Then
           If cboBar.ListIndex = 0 Then
                If ShowBarCountsForThisItem(grdCount.RowData(grdCount.Row)) Then
                    picBars.Left = grdCount.ColPos(grdCount.ColIndex("FullQty")) - 3900
                    If grdCount.RowPos(grdCount.Row) + picBars.Height < grdCount.Height - grdCount.RowHeight(0) Then
                        picBars.Top = grdCount.RowPos(grdCount.Row + 1)
                    Else
                        picBars.Top = grdCount.Height - picBars.Height - 30
                    End If
                    picBars.Visible = True
                End If
           End If
        End If
      End If
    End If

End Sub

Private Sub grdCount_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' ver 440
    picBars.Visible = False
    grdBars.Rows = 0

End Sub

Private Sub grdMenu_Click()
Dim sFolder As String

    If bEvaluation Then
        ' exit if evaluation and not correct menu selection
    
        If InStr("1#2#4#E#R#P#I#X", grdMenu.RowData(grdMenu.Row)) = 0 Then
            Exit Sub
        End If
    End If
    
    If MenuCheck(grdMenu.RowData(grdMenu.Row)) Then
    
        fraPrint.Visible = False
        fraDates.Visible = False
        fraStock.Visible = False
        fraDelivery.Visible = False
        fraSales.Visible = False
        picCash.Visible = False
        picSummary.Visible = False
        grdCount.Visible = False
        cmdPrint.Visible = True
        
        If btnSaveNote.Enabled Then
            btnSaveNote_Click
        End If
        
        SaveSummaryAndNote False
        ' btnSaveNote.Enabled = False
        
        SetupColours
        
        SetupShowPrevious False
        
        DoEvents
        
        Select Case grdMenu.RowData(grdMenu.Row)
        ' Select which row is clicked
        
            Case 1  ' Begin New Stock Take or Edit Stock Take Dates
             '                      Caption, PrintFrameVisible, gridvisible
            ' SetupFrameAndGrid "fraDates", 1, "", True, False
             
'             fraPrint.Visible = True
             
             fraDates.Left = (Me.Width - picStatus.Width - fraDates.Width) / 2 + picStatus.Width
             fraDates.Top = (Me.Height - picStatus.Top - fraDates.Height) / 2 + picStatus.Top
             
             fraDates.Visible = True

'VER 1.4
             If Not CountInProgress(lSelClientID) Then
             ' NEW STOCK TAKE
             
                lblStockTake.Caption = "Begin New Stock Take"
                chkEvaluation.Visible = True '

                If IsDate(labelFrom.Tag) And IsDate(labelTo.Tag) Then
                ' see if there are valid dates...

                    tedFrom.Text = Format(DateValue(labelTo.Tag) + 1, sDMY)
                    tedTo.Text = Format(Now, sDMY)
                    tedOn.Text = Format(DateAdd("d", 1, Now), sDMY)
                
                Else
                ' no valid dates so just show todays date

                    tedFrom.Text = Format(Now, sDMY)
                    tedTo.Text = Format(Now, sDMY)
                    tedOn.Text = Format(Now, sDMY)

                End If

             ElseIf IsDate(labelFrom.Tag) And IsDate(labelTo.Tag) Then
             ' Count in progress ... show dates saved
             
                lblStockTake.Caption = "Edit Dates of Stock Take"
                
                
                tedFrom.Text = Format(labelFrom.Tag, sDMY)
                tedTo.Text = Format(labelTo.Tag, sDMY)
             
                chkEvaluation.Visible = bEvaluation
                chkEvaluation.Value = Abs(bEvaluation)
                cmdCancel.Visible = Not bEvaluation ' only show cancel if its not an evaluation
             
             End If
             
             
             bSetFocus Me, "tedFrom"
             
            
' ver 520 ---------------------
            Case 2  ' Tablet / Count Sheets ' ver 500
             
             If MsgBox("Print Count Sheets?", vbDefaultButton1 + vbYesNo + vbQuestion, "Count Sheets") = vbYes Then
             ' ask output tablet file or print count sheets
             
                 cmdPrint.Visible = False
             
                 gbOk = PrintCountSheets(Replace(labelFrom.Tag, "/", "-"), lSelClientID, vbTrue, 1)
                 gbOk = PrintCountSheets(Replace(labelFrom.Tag, "/", "-"), lSelClientID, vbFalse, 1)
                 
                 LogMsg Me, "Count Sheets printed for " & Replace(lblClient.Tag, "_", " "), " From:" & labelFrom & " To:" & labelTo
                 
                 If MsgBox("Print Client Till Information?", vbDefaultButton1 + vbYesNo + vbQuestion, "Client Till Information") = vbYes Then
                    If PrintTillNotes(Replace(labelFrom.Tag, "/", "-"), lSelClientID, vbFalse, 1) Then
                    
                    Else
                        MsgBox "There are No Client Notes to Print"
                    End If
                 End If
    
                 labelTitle.Caption = ""
            End If
                 
             If MsgBox("Create Tablet Stock Count file?", vbDefaultButton1 + vbYesNo + vbQuestion, "Tablet Stock Count File") = vbYes Then
             ' ask output tablet file or print count sheets
                 
                Dim sDrive As String
                sDrive = voldrive("SWCount")
                ' ask is memory key loaded?
                
                ' search for memory key
                ' if found then pass to function
                
                gbOk = GenerateTabletOutputFile(sDrive, lSelClientID)
             
             
'                boResult = DismountVolume(sDrive)
'                CloseVolume (sDrive)
             
             
             End If
' ver 520 ---------------------
            
            
            Case 3  ' Print PLU Worksheets
             
             cmdPrint.Visible = False
             
             gbOk = PrintPLUCountSheet(Replace(labelFrom.Tag, "/", "-"), lSelClientID, 1)
             
             LogMsg Me, "PLU Count Sheet printed for " & Replace(lblClient.Tag, "_", " "), " From:" & labelFrom & " To:" & labelTo

             labelTitle.Caption = ""
             
            Case 4  ' Stock Count Entry
             
             LogMsg Me, "", ""
             
             Me.Picture = imgList.ListImages("stockblank").Picture
             
             SetupFrameAndGrid "fraStock", 4, "", True, True
             
             gbOk = InitStockEntry()
             ' clear the entry screen
             
             gbOk = ShowTotalStockCount()
             ' show stock to be counted for this client
             
             '----------------------------------------------------
             'Ver 440
             ' here we check to see if multiple bars have been set..
             
             If bMultipleBars Then
             ' case where multiple bars are present....
                
                gbOk = GetBarList()
                ' Get whatever bars are associated with this client
             
                ShowBarComboList True
                ' if there are multiple bars show them in the combo list
             
                cboBar.ListIndex = 0
                ' initially point to 'All Bars'
                
                InitEnableInput False
                ' disable inputting of any data to 'ALL Bars' since we only want
                ' figures going into real bars.
                
                grdCount.Row = 0
                
             Else
             ' Only one bar so... hide the combo box etc...
                
                ShowBarComboList False
                ' otherwise hide the combo box
             
                InitEnableInput True
                ' allow inputting of data straightoff since there is only one bar.

                If GetNextStockItem(lStkId) Then
                ' Get next blank stock rec to be filled in

                    bSetFocus Me, "txtFullQty"
                End If
             
             End If
             '------------------------------------------------------
            
            Case 5  ' Enter Deliveries
             
             LogMsg Me, "", ""
             
             Me.Picture = imgList.ListImages("deliveryblank").Picture
'             grdCount.BackColorBkg = &HCAE6D5
'             grdCount.BackColorFixed = &HCAE6D5
             
             SetupFrameAndGrid "fraDelivery", 5, "", True, True
             
             gbOk = InitDeliveryItem()
             
             gbOk = ShowDeliveries()
            
             grdCount.Row = 0
             ' Ver 3.0.7
             
             If GetNextDelivery(lNextDelID) Then
            
                bSetFocus Me, "txtQty"
            
             End If
             
            
            Case 6  ' Till Read or PLU SALES ENTRY
             
             LogMsg Me, "", ""
             
             Me.Picture = imgList.ListImages("salesblank").Picture
'             grdCount.BackColorBkg = &HC8E4F4
'             grdCount.BackColorFixed = &HC8E4F4
             
             SetupFrameAndGrid "fraSales", 6, "", True, True
             
             gbOk = InitTillItem()
             
             gbOk = ShowTillSales()
             
             If GetNextTillNo(lTillID) Then
             
                bSetFocus Me, "txtSales"
            
             End If
             
            Case 7  ' Enter Cash
             
             Me.Picture = imgList.ListImages("cashblank").Picture
             
             picCash.Top = (Me.Height - picSelect.Height + 1100 - picCash.Height) / 2
             picCash.Left = picStatus.Width + (Me.Width - picStatus.Width - picCash.Width) / 2


             
             LogMsg Me, "", ""
            
             picCash.Visible = True
             grdCount.Visible = False
             txtSurpluslbl.Text = "Cash €"
             txtOther.Text = "Other"
             
             gbOk = ShowCashTakings(lDatesID)
             

             If CountInProgress(lSelClientID) Then
             
                 gbOk = SetBoxColoursWhite()
                 
                 labelCalculatedSales = Format(GetCalculatedSales(lSelClientID), "0.00")
             
                 gbOk = getProjectedSales()

                 bSetFocus Me, "txtActual"
             End If
             
            '---------------------------------------------------------
            ' REPORTS
            
            Case "A"  ' Stock Analysis
            
             SetupFrameAndGrid "fraPrint", "A", grdMenu.Cell(flexcpTextDisplay, grdMenu.FindRow("A"), 1), True, True
             gbOk = RepStockAnalysis()
    
            Case "B"  ' Deliveries Report
             
             Me.Picture = imgList.ListImages("cashblank").Picture

             SetupFrameAndGrid "fraPrint", "B", grdMenu.Cell(flexcpTextDisplay, grdMenu.FindRow("B"), 1), True, True
             gbOk = RepDeliveries()
    
            Case "C"  ' Group Totals
            
             SetupFrameAndGrid "fraPrint", "C", grdMenu.Cell(flexcpTextDisplay, grdMenu.FindRow("C"), 1), True, True
             gbOk = RepGroupTotals()
            
            Case "D"  ' Stock/Till Reconciliation
             
             BlockAll True
             
             SetupFrameAndGrid "fraPrint", "D", grdMenu.Cell(flexcpTextDisplay, grdMenu.FindRow("D"), 1), True, True
             
             grdCount.Rows = 0
             DoEvents
             If cboShowPrevious = "" Then
                cboShowPrevious.ListIndex = 0
             Else
                gbOk = RepTillReconciliation(cboShowPrevious)
             End If
             
             SetupShowPrevious True
            
             BlockAll False
            
            Case "E"  ' Closing Stock & Valuations
             
             SetupFrameAndGrid "fraPrint", "E", grdMenu.Cell(flexcpTextDisplay, grdMenu.FindRow("E"), 1), True, True
             gbOk = RepClosingStock()
             
            Case "F"    ' Profit Discrepancy
            
             SetupFrameAndGrid "fraPrint", "F", grdMenu.Cell(flexcpTextDisplay, grdMenu.FindRow("F"), 1), True, True
             gbOk = RepProfitDiscrepance()
            
            Case "G"    ' Summary Analysis  NOW CALLED: Product Deficit
             
             SetupFrameAndGrid "fraPrint", "G", grdMenu.Cell(flexcpTextDisplay, grdMenu.FindRow("G"), 1), True, True
             gbOk = RepSummaryAnalysis()

             txtNote.Text = GetSummaryDetails(lDatesID, True)
            
             gbOk = totalSelected()
            
            Case "R"    ' generate Reports
            
             cmdPrint.Visible = False

             If ConfirmClientName() Then
             
                  If MsgBox("Are you sure you want to generate reports for " & Replace(lblClient.Tag, "_", " "), vbYesNo, "Generate Reports") = vbYes Then
                
'                    imgReport.Visible = False
                    
                    If MsgBox("Please ensure Microsoft Word is not open during report generation", vbDefaultButton1 + vbOKCancel + vbInformation, "Microsoft Word Check") = vbOK Then
                    
                        gbOk = TerminateWINWORD()
                        
                        BlockAll True
                        ' Block all other commands
                        
                        LogMsg Me, "Generating Reports for " & Replace(lblClient.Tag, "_", " "), " From:" & labelFrom & " To:" & labelTo
                        
                        ' Report I (INVOICE)
                        If Not InvoiceCreated(lDatesID) Then
                        ' Show it if its there already
             
'                            If MsgBox("Create Invoice Now?", vbDefaultButton1 + vbQuestion + vbYesNo, "Invoice not created yet") = vbYes Then
'                                LogMsg Me, "Generating Invoice for " & Replace(lblClient.Tag, "_", " "), " From:" & labelFrom & " To:" & labelTo
                           GoCreateInvoice
                           LogMsg Me, "Invoice Generated for " & Replace(lblClient.Tag, "_", " "), " From:" & labelFrom & " To:" & labelTo
                            
'                            End If
                            
                        End If
                        
                        bHourGlass True
                        
                        
                        SetupFrameAndGrid "fraPrint", "G", grdMenu.Cell(flexcpTextDisplay, grdMenu.FindRow("G"), 1), True, False
                        
                        gbOk = CreateClientFolder(sFolder, lblClient.Tag, Replace(labelTo.Tag, "/", "-"))
                        ' Check is Folder created  - if not create it
                        ' i.e. \StockWatch\Client Name\
                              
' GoTo SummaryRep
                        
                        
                        If Not bEvaluation Then

                            ' REPORT A
                            If RepStockAnalysis() Then
                                gbOk = CreateReport(sFolder, grdMenu.Cell(flexcpTextDisplay, grdMenu.FindRow("A"), 1), 40, 2)
                                LogMsg Me, "Stock Analysis Report Generated for " & Replace(lblClient.Tag, "_", " "), " From:" & labelFrom & " To:" & labelTo
                            End If
                        
                            ' REPORT B
                            If RepDeliveries() Then
                                gbOk = CreateReport(sFolder, grdMenu.Cell(flexcpTextDisplay, grdMenu.FindRow("B"), 1), iRepLines, 1)
                                LogMsg Me, "Purchases Report Generated for " & Replace(lblClient.Tag, "_", " "), " From:" & labelFrom & " To:" & labelTo
                            End If
                        
                            ' REPORT C
                            If RepGroupTotals() Then
                                gbOk = CreateReport(sFolder, grdMenu.Cell(flexcpTextDisplay, grdMenu.FindRow("C"), 1), iRepLines, 2)
                                LogMsg Me, "Group totals Report Generated for " & Replace(lblClient.Tag, "_", " "), " From:" & labelFrom & " To:" & labelTo
                            End If
                        
                            ' REPORT D
                            If RepTillReconciliation("4") Then
                                gbOk = CreateReport(sFolder, grdMenu.Cell(flexcpTextDisplay, grdMenu.FindRow("D"), 1), iRepLines, 2)
                                LogMsg Me, "Stock/Till Reconciliation Report Generated for " & Replace(lblClient.Tag, "_", " "), " From:" & labelFrom & " To:" & labelTo
                            End If
                            ' default to showing 4 previous difference records
                        End If
                        
                        ' REPORT E
                        If RepClosingStock() Then
                            gbOk = CreateReport(sFolder, grdMenu.Cell(flexcpTextDisplay, grdMenu.FindRow("E"), 1), iRepLines, 2)
                            LogMsg Me, "Closing Stock Report Generated for " & Replace(lblClient.Tag, "_", " "), " From:" & labelFrom & " To:" & labelTo
                        End If
                        ' Only report neded for an evaluation
                        
                        If Not bEvaluation Then
                        
                            ' REPORT F
                            If RepProfitDiscrepance() Then
                                gbOk = CreateReport(sFolder, grdMenu.Cell(flexcpTextDisplay, grdMenu.FindRow("F"), 1), 46, 1)
                                LogMsg Me, "Profit Discrepancy Report Generated for " & Replace(lblClient.Tag, "_", " "), " From:" & labelFrom & " To:" & labelTo
                            End If
                

                            ' REPORT G  ' NOW PRODUCT DEFICIT
                            If Not bEvaluation And RepSummaryAnalysis() Then
                        
                                txtNote.Text = GetSummaryDetails(lDatesID, True)
                        
                                gbOk = totalSelected()
                            
                            
                                ShowSelected
                                
                                RemoveHiddenRows
                            
                                'ver 310
                                ' make sure only selected products are shown
                                ' this ensures only one page will be printed.
                            
                                gbOk = CreateReport(sFolder, grdMenu.Cell(flexcpTextDisplay, grdMenu.FindRow("G"), 1), 29, 1)
                                LogMsg Me, "Product Deficit Report Generated for " & Replace(lblClient.Tag, "_", " "), " From:" & labelFrom & " To:" & labelTo
                            End If
                        
                        End If
                        
'Ver441
                        gbOk = GenerateReportCover(lblClient.Tag, Replace(labelTo.Tag, "/", "-"))
                        
                        LogMsg Me, "Reports complete for " & Replace(lblClient.Tag, "_", " "), " From:" & labelFrom & " To:" & labelTo
                    
                    End If
                
                End If
             End If
             
             
             BlockAll False
             ' allow other menu commands
             
             imgReport.Visible = False
             picSummary.Visible = False
             grdCount.Visible = False
             ' tidy up!
             
             bHourGlass False

            Case "P"  ' Print Reports
             
             cmdPrint.Visible = False
             
             If ConfirmClientName() Then
                 
                frmPrintSelect.Show vbModal
                 
             End If
            
            Case "I"    ' Invoice
             
             ' see if invoice created already and display if it has been
            
             If InvoiceCreated(lDatesID) Then
             ' Show it if its there already
             
                gbOk = GetReport(lDatesID, "Invoice")
             
             Else
             ' otherwise show invoice create form
             
                 GoCreateInvoice
             End If
             
            
            Case "X"    ' Close Stock Take
             
             If bEvaluation Then
             
               If MsgBox("Close Stock Valuation for " & Replace(lblClient.Tag, "_", " "), vbDefaultButton2 + vbYesNo + vbQuestion, "End Valuation") = vbYes Then
         
                 LogMsg Me, "Closing Valuation for " & Replace(lblClient.Tag, "_", " "), " on:" & labelOn
                 
                 If Not InvoiceCreated(lDatesID) Then
                 ' Show it if its there already
             
                     GoCreateInvoice
        
                 End If
 
                 If RestoreLastCountFigures(lSelClientID) Then

                    gbOk = SetEvaluationFlag(lSelClientID, False)
                    ' clear the evaluation flag
                    
                    gbOk = ClearValuation(lDatesID)
                    ' Clear the Inprogress Flag and id
                 
                    bEvaluation = False
                    chkEvaluation.Value = 0
                    
                    If SetSymbolMsgMenu(lSelClientID) Then
    
                        gbOk = SetBeginEndStockTakeButtons(lSelClientID)
                                
                        ShowMenu vbTrue
        
                    End If
                     
                 End If
              End If
                 
             Else   ' Close of regular stock take
             
               If MsgBox("Close Stock Take for " & Replace(lblClient.Tag, "_", " "), vbDefaultButton2 + vbYesNo + vbQuestion, "Save and End Stock Take") = vbYes Then
         
                LogMsg Me, "Closing Stock Take for " & Replace(lblClient.Tag, "_", " "), " From:" & labelFrom & " To:" & labelTo
                
                If Not InvoiceCreated(lDatesID) Then
                ' Show it if its there already
             
                     GoCreateInvoice
        
                End If
                
                ' ver 3.0.6 (7) if cancel during invoice create could
                ' close stocktake without creating an invoice
                
                If InvoiceCreated(lDatesID) Then
                ' make sure its now created...
                
                    gbOk = ClearInProgress(lDatesID)
                    ' Clear the Inprogress Flag and id
                 
                    If GetClientDate(lSelClientID, lDatesID) Then
                    ' dates of audit
        
                        If SetSymbolMsgMenu(lSelClientID) Then
                        ' Audit symbol
                        ' In Progress message in red
                        ' modified Menu
    
                            gbOk = SetBeginEndStockTakeButtons(lSelClientID)
                                
                            ShowMenu vbTrue
        
                        End If
                    
                    End If
    
                    gbOk = SetBeginEndStockTakeButtons(lSelClientID)
                
                End If
               End If
             End If
             
        End Select
    
    
    End If

End Sub


Private Sub grdMenu_KeyPress(KeyAscii As Integer)
Dim iRow As Integer
    
    If KeyAscii = 13 Then
        grdMenu_Click

    Else

        iRow = grdMenu.FindRow(UCase(Chr$(KeyAscii)), , 0)
        
        If iRow > -1 Then
            grdMenu.Row = iRow
            grdMenu_Click
        End If

    End If
    
End Sub

Private Sub grdMenu_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyUp Then
        If grdMenu.Row = 0 Then
            grdMenu.Row = 1
        ElseIf grdMenu.Row = 4 Then
            grdMenu.Row = 3
        ElseIf grdMenu.Row = 10 Then
            grdMenu.Row = 9
        ElseIf grdMenu.Row = 18 Then
            grdMenu.Row = 17
'        ElseIf grdMenu.Row = 19 Then
'            grdMenu.Row = 18
'        ElseIf grdMenu.Row = 20 Then
'            grdMenu.Row = 19
        End If
    
    ElseIf KeyCode = vbKeyDown Then
        If grdMenu.Row = 4 Then
            grdMenu.Row = 5
        ElseIf grdMenu.Row = 10 Then
            grdMenu.Row = 11
        ElseIf grdMenu.Row = 18 Then
            grdMenu.Row = 19
'        ElseIf grdMenu.Row = 18 Then
'            grdMenu.Row = 19
'        ElseIf grdMenu.Row = 19 Then
'            grdMenu.Row = 20
        End If
    End If

End Sub


Private Sub lblClient_Click()

    cmdSelectClient_Click

End Sub


Public Function ShowActiveClientList()
Dim rs As Recordset
Dim sLastClient As String

    On Error GoTo ErrorHandler

    bHourGlass True
    
    grdClients.Rows = 0
    
'VER1.4
    Set rs = SWdb.OpenRecordset("Select txtNAme, rtfAddress, tblDates.InProgress, tblClients.ID FROM tblClients LEFT JOIN tblDates ON tblClients.ID = tblDates.ClientID WHERE chkActive = true ORDER BY txtName, tblDates.InProgress", dbOpenSnapshot)
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do
            If sLastClient <> rs("txtName") Then
            
                grdClients.AddItem vbTab & rs("txtName") & vbTab & _
                                Replace(rs("rtfAddress"), vbCrLf, " ")
                
                sLastClient = rs("txtName")
    
                grdClients.RowData(grdClients.Rows - 1) = rs("ID") + 0
                
                If rs("InProgress") Then
                
                    grdClients.Cell(flexcpPicture, grdClients.Rows - 1, 0) = imgList.ListImages("inp").Picture
                    grdClients.Cell(flexcpData, grdClients.Rows - 1, 0) = True
                    
                End If
            
            End If
            
            rs.MoveNext
        Loop While Not rs.EOF
    End If
    
    If grdClients.Rows > 30 Then
    
        grdClients.Height = (grdClients.RowHeight(0) * 30) + 60
    
    Else
        grdClients.Height = (grdClients.RowHeight(0) * grdClients.Rows) + 60
    End If
    
    ShowActiveClientList = True

CleanExit:
    
    bHourGlass False
    
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
'    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("ShowActiveClientList") Then Resume 0
    Resume CleanExit

End Function



Private Sub MyButton2_Click()

End Sub

Private Sub tedFrom_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub tedFrom_LostFocus()
    lblFrom.ForeColor = sBlack

End Sub

Private Sub tedOn_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub tedOn_LostFocus()
    labelOn.ForeColor = sBlack

End Sub

Private Sub tedTo_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Public Function GetPLUGroupID(lCLId As Long, iGrpNo As Integer)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("SELECT * FROM tblPLUGroup WHERE ClientID = " & Trim$(lCLId) & " AND txtGroupNumber = " & Trim$(iGrpNo), dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        GetPLUGroupID = rs("ID") + 0
    End If
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
'    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetPLUGroupID") Then Resume 0
    Resume CleanExit

End Function

Private Sub tedTo_LostFocus()
    lblTo.ForeColor = sBlack

End Sub

Private Sub timXfer_Timer()

End Sub

'''
'''Private Sub timXfer_Timer()
''''Dim sTemp As String
'''
'''    ' try every 1/2 hour to send files and emails
'''
''''    timXfer.Enabled = False
'''
'''    ' using timer to invoke this here to
'''    ' allow sw to start up fully first
'''
'''    If iMins > 30 Then
'''
'''        If Not gbTestEnvironment Then
'''            gbOk = SendInvoiceBySMTP()
'''        End If
'''
'''
'''        iMins = 0
'''
'''    Else
'''        iMins = iMins + 1
'''        ' count up minutes
'''    End If
'''
'''
'''
'''End Sub

Private Sub txtActual_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtActual_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
        ' kill beep

    ElseIf KeyAscii <> 46 Then
        
        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
        
    ElseIf InStr(txtActual, ".") > 0 Then
            
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtActual_LostFocus()
    lblActual.ForeColor = sBlack

End Sub


'''Private Sub txtCode_GotFocus()
'''
'''    txtCode.BackColor = sWhite
'''    InitStockEntry
'''
'''End Sub
'''
'''Private Sub txtCode_KeyUp(KeyCode As Integer, Shift As Integer)
'''
'''    If Val(txtCode) > 999 Then
'''        lStkId = GetStockID(Val(txtCode))
'''        gbOk = GetStockItem(lStkId)
'''        bSetFocus Me, "txtFullQty"
'''
'''    End If
'''
'''
'''End Sub
'''
'''Private Sub txtCode_LostFocus()
'''
'''    txtCode.Text = ""
'''    txtCode.BackColor = &HDECCB8
'''
'''End Sub

Private Sub txtComplimentary_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtComplimentary_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            KeyAscii = 0
        ' kill beep

    ElseIf KeyAscii <> 46 Then
        
        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
        
    ElseIf InStr(txtComplimentary, ".") > 0 Then
            
        KeyAscii = 0
    End If

End Sub

Private Sub txtComplimentary_KeyUp(KeyCode As Integer, Shift As Integer)
    gbOk = getProjectedSales()

End Sub

Private Sub txtComplimentary_LostFocus()
    lblComplimentary.ForeColor = sBlack

End Sub


Private Sub txtDelCost_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtDelCost_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        ' kill beep
    End If

End Sub

Private Sub txtDelCost_LostFocus()
    lblDelCost.ForeColor = sBlack

End Sub

Private Sub txtDelOther_Change()
    SetEditButton True

End Sub

Private Sub txtDelOther_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtDelOther_KeyPress(KeyAscii As Integer)

'    cmdAddNew.Enabled = True


End Sub

Private Sub txtDelOther_LostFocus()
    lblDelivery.ForeColor = sBlack

End Sub

Private Sub txtFree_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtFree_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Then
        bSetFocus Me, "grdCount"
    
    ElseIf KeyCode = vbKeyUp Then
        bSetFocus Me, "txtQty"
    
    End If
    
End Sub

Private Sub txtFree_KeyPress(KeyAscii As Integer)
    cmdDelSave.Enabled = True
'    cmdAddNew.Enabled = True

    If KeyAscii = 13 Then
            
        KeyAscii = 0
        ' kill beep
        
        cmdDelSave_Click
        Pause 200
        
    Else
    
        KeyAscii = CharOk(KeyAscii, 0, "") ' 0 = no only, 1 = char only, 2 = both
    
    End If

End Sub

Private Sub txtFree_LostFocus()
    lblFree.ForeColor = sBlack

End Sub

Private Sub txtFullQty_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtFullQty_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Then
            
        If grdCount.Row < grdCount.Rows - 1 Then
            If grdCount.RowData(grdCount.Row + 1) <> Empty Then
                grdCount.Row = grdCount.Row + 1
                bSetFocus Me, "grdCount"
                grdCount_Click
                If grdCount.Row > 22 Then grdCount.TopRow = grdCount.TopRow + 1
                bSetFocus Me, "txtfullqty"
            
            Else
                If grdCount.Row < grdCount.Rows - 2 Then
                    If grdCount.RowData(grdCount.Row + 2) <> Empty Then
                        grdCount.Row = grdCount.Row + 2
                        bSetFocus Me, "grdCount"
                        grdCount_Click
                        If grdCount.Row > 22 Then grdCount.TopRow = grdCount.TopRow + 2
                        bSetFocus Me, "txtfullqty"
                    
                    End If
                End If
            End If
        End If
            
    ElseIf KeyCode = vbKeyUp Then
    
        If grdCount.Row > 1 Then
            If grdCount.RowData(grdCount.Row - 1) <> Empty Then
                grdCount.Row = grdCount.Row - 1
                bSetFocus Me, "grdCount"
                grdCount_Click
                If grdCount.Row > 20 Then grdCount.TopRow = grdCount.TopRow - 1
                bSetFocus Me, "txtfullqty"
            Else
                If grdCount.Row > 2 Then
                    grdCount.Row = grdCount.Row - 2
                    bSetFocus Me, "grdCount"
                    grdCount_Click
                    If grdCount.Row > 21 Then grdCount.TopRow = grdCount.TopRow - 2
                    bSetFocus Me, "txtfullqty"
                End If
            End If
        End If
    
    Else
        
        cmdSaveStock.Enabled = True
    End If

End Sub

Private Sub txtFullQty_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        KeyAscii = 0
        ' kill beep
        
        If bAddIt Then
        
            bAddIt = False
            txtFullQty = AddUp(Trim$(txtFullQty.Text))
            txtFullQty.Width = txtOpen.Width
            txtFullQty.ForeColor = sBlack
            If txtOpen.Visible Then
                bSetFocus Me, "txtOpen"
            Else
                ' ver 440
                'bSetFocus Me, "cmdSaveStock"
                cmdSaveStock_Click
            End If
            
        ElseIf lblFullQty.Tag = "False" Then
            
            cmdSaveStock.Enabled = True
            
            If txtFullQty = "" Then
                txtFullQty = "0"
            End If
            
            cmdSaveStock_Click
            Pause 200
        
        End If
         
    ElseIf KeyAscii = 43 Then
    ' +
            
        txtFullQty.SelStart = Len(txtFullQty)
    
        txtFullQty.Width = SetBoxWidth(txtFullQty)
        
        txtFullQty.ForeColor = vbBlue
        bAddIt = True
        
''Ver 2.1.0
'    ElseIf txtOpen.Visible = True Then
'
'        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
'        txtFullQty.Width = SetBoxWidth(txtFullQty)
'
'        cmdSaveStock.Enabled = True
'
''==========
    
    Else
    
        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
        
        txtFullQty.Width = SetBoxWidth(txtFullQty)
    
        cmdSaveStock.Enabled = True
    
    End If


End Sub

Private Sub txtFullQty_KeyUp(KeyCode As Integer, Shift As Integer)
    
    txtFullQty.Width = SetBoxWidth(txtFullQty)

    If (KeyCode = 187 And Shift = 1) Or KeyCode = vbKeyShift Or InStr(txtFullQty, "+") <> 0 Then
    
'    Else
'        txtFullQty.SelStart = 0
'        txtFullQty.SelLength = Len(txtFullQty)
    End If
    
End Sub

Private Sub txtFullQty_LostFocus()
    
    bAddIt = False
    lblFullQty.ForeColor = sBlack

End Sub



Private Sub txtGlass_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtGlass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        KeyAscii = 0
        ' kill beep
        
        If bDualPrice Then
            bSetFocus Me, "txtSalesDP"
        Else
            cmdTillSave_Click
            Pause 200
        End If
    Else
        KeyAscii = CharOk(KeyAscii, 0, "") ' 0 = no only, 1 = char only, 2 = both

        If KeyAscii <> 0 Then
            cmdTillSave.Enabled = True
        End If
        
    End If

End Sub

Private Sub txtGlass_LostFocus()
    lblGlass.ForeColor = sBlack

End Sub

Private Sub txtGlassDP_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtGlassDP_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        
        cmdTillSave_Click
        Pause 200
    Else
        KeyAscii = CharOk(KeyAscii, 0, "") ' 0 = no only, 1 = char only, 2 = both

        If KeyAscii <> 0 Then
            cmdTillSave.Enabled = True
        End If
        
    End If
    


End Sub

Private Sub txtGlassDP_LostFocus()
    lblGlassDP.ForeColor = sBlack

End Sub

Private Sub txtKitchen_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtKitchen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            KeyAscii = 0
        ' kill beep

    ElseIf KeyAscii <> 46 Then
        
        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
        
    ElseIf InStr(txtKitchen, ".") > 0 Then
            
        KeyAscii = 0
    End If

End Sub

Private Sub txtKitchen_KeyUp(KeyCode As Integer, Shift As Integer)
    gbOk = getProjectedSales()

End Sub

Private Sub txtKitchen_LostFocus()
    lblKitchen.ForeColor = sBlack

End Sub

Private Sub txtNote_KeyPress(KeyAscii As Integer)
        
    If KeyAscii = 124 Then KeyAscii = 0
    If KeyAscii = 126 Then KeyAscii = 0

    'If KeyAscii <> 0 Then btnSaveNote.Enabled = True
        
    If KeyAscii <> 0 Then SaveSummaryAndNote True
End Sub

Private Sub txtOffLicense_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtOffLicense_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            KeyAscii = 0
        ' kill beep

    ElseIf KeyAscii <> 46 Then
        
        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
        
    ElseIf InStr(txtOffLicense, ".") > 0 Then
            
        KeyAscii = 0
    End If

End Sub

Private Sub txtOffLicense_KeyUp(KeyCode As Integer, Shift As Integer)
    gbOk = getProjectedSales()

End Sub

Private Sub txtOffLicense_LostFocus()
    lblOffLicense.ForeColor = sBlack

End Sub

Private Sub txtOpen_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtOpen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        bSetFocus Me, "txtWeight"
    ElseIf KeyCode = vbKeyUp Then
        bSetFocus Me, "txtFullQty"
    End If
    
End Sub

Private Sub txtOpen_KeyPress(KeyAscii As Integer)

'    cmdSaveStock.Enabled = True
'
'    KeyAscii = CharOk(KeyAscii, 0, "") ' 0 = no only, 1 = char only, 2 = both
'
'    If KeyAscii = 13 Then KeyAscii = 0
'    ' kill beep

    If KeyAscii = 13 Then
        
        KeyAscii = 0
        ' kill beep
        
        If bAddIt Then
        
            bAddIt = False
            txtOpen = AddUp(Trim$(txtOpen.Text))
            txtOpen.Width = iBoxWidth
            txtOpen.ForeColor = sBlack
            bSetFocus Me, "txtWeight"
        
        End If
         
    ElseIf KeyAscii = 43 Then
    ' +
            
        txtOpen.SelStart = Len(txtOpen)
    
        txtOpen.Width = SetBoxWidth(txtOpen)
        
        txtOpen.ForeColor = vbBlue
        bAddIt = True
        
    Else
    
        KeyAscii = CharOk(KeyAscii, 0, "") ' 0 = no only, 1 = char only, 2 = both
        
        txtOpen.Width = SetBoxWidth(txtOpen)
    
        cmdSaveStock.Enabled = True
    
    End If



End Sub


Private Sub txtOpen_KeyUp(KeyCode As Integer, Shift As Integer)
    txtOpen.Width = SetBoxWidth(txtOpen)

End Sub

Private Sub txtOpen_LostFocus()
    
    bAddIt = False
    lblOpen.ForeColor = sBlack

End Sub

Private Sub txtOther_GotFocus()
    gbOk = bSetupControl(Me)
    txtOtherLbl.ForeColor = &HFF0000

End Sub

Private Sub txtOther_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 46 Then
        
        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
        
        If KeyAscii = 13 Then
            
            KeyAscii = 0
            ' kill beep
            
        End If
    
    ElseIf InStr(txtOther, ".") > 0 Then
            
        KeyAscii = 0
    
    End If

End Sub

Private Sub txtOther_KeyUp(KeyCode As Integer, Shift As Integer)
    gbOk = getProjectedSales()

End Sub

Private Sub txtOther_LostFocus()
    txtOtherLbl.ForeColor = sBlack

End Sub

Private Sub txtOtherLbl_GotFocus()

    txtOtherLbl.SelStart = 0
    txtOtherLbl.SelLength = Len(txtOtherLbl.Text)

    txtOtherLbl.BackColor = sWhite
    txtOtherLbl.ForeColor = sBlack

End Sub

Private Sub txtOtherLbl_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then KeyAscii = 0
        ' kill beep


End Sub

Private Sub txtOtherLbl_LostFocus()
    txtOtherLbl.ForeColor = sBlack
    txtOtherLbl.BackColor = sRed
End Sub
Private Sub txtSurplusLbl_GotFocus()

    txtSurpluslbl.SelStart = 0
    txtSurpluslbl.SelLength = Len(txtSurpluslbl.Text)

    txtSurpluslbl.BackColor = sWhite
    txtSurpluslbl.ForeColor = &HFF0000
End Sub

Private Sub txtSurplusLbl_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then KeyAscii = 0
        ' kill beep


End Sub

Private Sub txtSurplusLbl_LostFocus()
    txtSurpluslbl.ForeColor = sBlack
    txtSurpluslbl.BackColor = sRed
End Sub

Private Sub txtOverRings_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtOverRings_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            KeyAscii = 0
        ' kill beep

    ElseIf KeyAscii <> 46 Then
        
        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
        
    ElseIf InStr(txtOverRings, ".") > 0 Then
            
        KeyAscii = 0
    End If

End Sub

Private Sub txtOverRings_KeyUp(KeyCode As Integer, Shift As Integer)
    gbOk = getProjectedSales()

End Sub

Private Sub txtOverRings_LostFocus()
    lblOverRings.ForeColor = sBlack

End Sub


Private Sub txtPromotions_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtPromotions_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            KeyAscii = 0
        ' kill beep

    ElseIf KeyAscii <> 46 Then
        
        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
        
    ElseIf InStr(txtPromotions, ".") > 0 Then
            
        KeyAscii = 0
    End If

End Sub

Private Sub txtPromotions_KeyUp(KeyCode As Integer, Shift As Integer)
    gbOk = getProjectedSales()

End Sub

Private Sub txtPromotions_LostFocus()
    lblPromotions.ForeColor = sBlack

End Sub

Private Sub txtQty_Change()
    SetEditButton True

End Sub

Private Sub txtQty_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
    '    cmdDelSave_Click
    
    ElseIf KeyCode = vbKeyDown Then
            
        If grdCount.Row < grdCount.Rows - 1 Then
            If grdCount.RowData(grdCount.Row + 1) <> Empty Then
                grdCount.Row = grdCount.Row + 1
                bSetFocus Me, "grdCount"
                grdCount_Click
                If grdCount.Row > 22 Then grdCount.TopRow = grdCount.TopRow + 1
                bSetFocus Me, "txtQty"
            
            End If
        End If
            
    ElseIf KeyCode = vbKeyUp Then
    
        If grdCount.Row > 1 Then
            If grdCount.RowData(grdCount.Row - 1) <> Empty Then
                grdCount.Row = grdCount.Row - 1
                bSetFocus Me, "grdCount"
                grdCount_Click
                If grdCount.Row > 21 Then grdCount.TopRow = grdCount.TopRow - 1
                bSetFocus Me, "txtQty"
            End If
        End If
    End If
    
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            
        KeyAscii = 0
        ' kill beep
        
        cmdDelSave_Click
        Pause 200
        
    ElseIf cboDelivery.ListIndex = 0 Then
    
        KeyAscii = CharOk(KeyAscii, 0, "") ' 0 = no only, 1 = char only, 2 = both
    
    Else
    
        KeyAscii = CharOk(KeyAscii, 0, "-.") ' 0 = no only, 1 = char only, 2 = both
    
    End If
    
End Sub

Private Sub txtQty_KeyUp(KeyCode As Integer, Shift As Integer)


'    txtQty.SelStart = 0
'    txtQty.SelLength = Len(txtQty)

    
    If cboDelivery.ListIndex = 1 Then
    ' Other selected
    
        If Val(txtQty.Text) < 0 Then
            
            cmdDelSave.Enabled = True
        
        Else
            cmdDelSave.Enabled = False
        
        End If
    
    End If
    
End Sub

Private Sub txtQty_LostFocus()
    lblQty.ForeColor = sBlack

End Sub

Private Sub txtSales_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtSales_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Then
            
        If grdCount.Row < grdCount.Rows - 1 Then
            grdCount.Row = grdCount.Row + 1
            
            bSetFocus Me, "grdCount"
            grdCount_Click
            If grdCount.Row > 22 Then grdCount.TopRow = grdCount.TopRow + 1
            bSetFocus Me, "txtSales"
        
        End If
            
    ElseIf KeyCode = vbKeyUp Then
    
        If grdCount.Row > 1 Then
            grdCount.Row = grdCount.Row - 1
            bSetFocus Me, "grdCount"
            grdCount_Click
            If grdCount.Row > 21 Then grdCount.TopRow = grdCount.TopRow - 1
            bSetFocus Me, "txtSales"
        
        End If
    
    End If
    
End Sub

Private Sub txtSales_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        
        KeyAscii = 0
        ' kill beep
        
        If txtGlass.Visible Then
            bSetFocus Me, "txtGlass"
        ElseIf bDualPrice Then
            bSetFocus Me, "txtSalesDP"
        Else
            cmdTillSave_Click
            Pause 200
        End If
    
    Else
        KeyAscii = CharOk(KeyAscii, 0, "") ' 0 = no only, 1 = char only, 2 = both

        If KeyAscii <> 0 Then
            cmdTillSave.Enabled = True
        End If
        
    End If
    

End Sub

Private Sub txtSales_LostFocus()
    lblSales.ForeColor = sBlack

End Sub




Private Sub txtSalesDP_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtSalesDP_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        KeyAscii = 0
        ' kill beep
        
        
        ' ver 549 .. also check if txtGlassDP is visible
        
        If iGlass > 0 And txtGlassDP.Visible = True Then
            bSetFocus Me, "txtGlassDP"
        Else
            cmdTillSave_Click
            Pause 200
        End If
        
    Else
        KeyAscii = CharOk(KeyAscii, 0, "") ' 0 = no only, 1 = char only, 2 = both

        If KeyAscii <> 0 Then
            cmdTillSave.Enabled = True
        End If
        
    End If

End Sub

Private Sub txtSalesDP_LostFocus()
    lblSalesDP.ForeColor = sBlack

End Sub

Private Sub txtStaff_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtStaff_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
        ' kill beep

    ElseIf KeyAscii <> 46 Then
        
        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
        
    ElseIf InStr(txtStaff, ".") > 0 Then
            
        KeyAscii = 0
    End If

End Sub

Private Sub txtStaff_KeyUp(KeyCode As Integer, Shift As Integer)

    gbOk = getProjectedSales()
    

End Sub

Private Sub txtStaff_LostFocus()
    lblStaff.ForeColor = sBlack

End Sub

Public Function CountInProgress(lID As Long) As Boolean
Dim rs As Recordset
    
    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("SELECT InProgress FROM tblDates WHERE ClientID = " & Trim$(lID) & " AND InProgress = true", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        
        CountInProgress = True
    
    End If
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks

    If Not rs Is Nothing Then Set rs = Nothing

    Exit Function

ErrorHandler:
    If CheckDBError("CountInProgress") Then Resume 0
    Resume CleanExit



End Function
                    
Public Function ShowClient(lCLId As Long) As Boolean
Dim rs As Recordset
            
    On Error GoTo ErrorHandler
    
    If lCLId > 0 Then
    ' as long there's a client to show...
    
        Set rs = SWdb.OpenRecordset("tblClients")
        rs.Index = "PrimaryKey"
        rs.Seek "=", lCLId
        If Not rs.NoMatch Then
                
            lblClient.Caption = Replace(rs("txtName") & vbCrLf & Replace(rs("rtfAddress"), vbCrLf, " "), "&", "&&")
            lblClient.Tag = Replace(Trim$(rs("txtName")), " ", "_")
            ' show name and address
            
            bDualPrice = rs("chkDualPricing")
            bMultipleBars = rs("chkMultipleBars")
        
        End If
        
        rs.Close
    
        ShowClient = True
    
    End If
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
'
    
    Exit Function

ErrorHandler:
    If CheckDBError("ShowSelectedClient") Then Resume 0
    Resume CleanExit


End Function

Public Function GetCountDates(lCLId As Long, bNextCountDates As Boolean)
Dim rs As Recordset
            
    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("SELECT * FROM tblDates WHERE ClientID = " & Trim$(lCLId) & " ORDER BY To", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
            
        rs.MoveFirst
            
        If bNextCountDates Then
        ' For next count dates, Begin count on date following last count
        ' and default to today for the 'to' date.
        
            tedFrom = Format(rs("To") + 1, sDMY)
            tedTo = Format(Now, sDMY)
        
        Else
        ' For last count just show from and to dates of last count
        
            tedFrom = Format(rs("From"), sDMY)
            tedTo = Format(rs("To"), sDMY)
        End If
        
    Else
    ' New Customer show todays date
    
        tedFrom = Format(Now, sDMY)
        tedTo = Format(Now, sDMY)
    
    End If
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
'
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetCountDates") Then Resume 0
    Resume CleanExit



End Function


Public Function StartCount(lID As Long, dtFrom As Date, dtTo As Date, dtOn As Date)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblDates")
    rs.Index = "PrimaryKey"
    
    If lID = 0 Then
    ' New count
        rs.AddNew
        rs("ClientID") = lSelClientID
        rs("From") = dtFrom
        rs("To") = dtTo
        rs("On") = dtOn
        rs("InProgress") = True
        rs.Update
    
    Else
    ' editing dates
        rs.Seek "=", lID
        If Not rs.NoMatch Then
            rs.Edit
            rs("From") = dtFrom
            rs("To") = dtTo
            rs("On") = dtOn
            rs.Update
        End If
    
    End If
    
    StartCount = True
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
    
        If CheckDBError("StartCount") Then Resume 0
        Resume CleanExit
    End If
    
End Function


Public Sub SetUpMenuItemsColoursButtons()
Dim iRow As Integer

    ' set row datas for convenience
    
    grdMenu.RowData(1) = 1
    grdMenu.RowData(2) = 2
    grdMenu.RowData(3) = 3
    
    grdMenu.RowData(5) = 4
    grdMenu.RowData(6) = 5
    grdMenu.RowData(7) = 6
    grdMenu.RowData(8) = 7
    
    grdMenu.RowData(10) = "A"
    grdMenu.RowData(11) = "B"
    grdMenu.RowData(12) = "C"
    grdMenu.RowData(13) = "D"
    grdMenu.RowData(14) = "E"
    grdMenu.RowData(15) = "F"
    grdMenu.RowData(16) = "G"
    
    grdMenu.RowData(18) = "R"
    grdMenu.RowData(19) = "P"
    grdMenu.RowData(20) = "I"
    grdMenu.RowData(21) = "X"


    grdMenu.Cell(flexcpBackColor, 0, 0, 0, 1) = sDarkPurple
    grdMenu.Cell(flexcpBackColor, 4, 0, 4, 1) = sDarkPurple
    grdMenu.Cell(flexcpBackColor, 9, 0, 9, 1) = sDarkPurple
    grdMenu.Cell(flexcpBackColor, 17, 0, 17, 1) = sDarkPurple

    For iRow = 1 To 3
        grdMenu.Cell(flexcpPicture, iRow, 0) = imgList.ListImages("button").Picture
    Next
    For iRow = 5 To 8
        grdMenu.Cell(flexcpPicture, iRow, 0) = imgList.ListImages("button").Picture
    Next
    For iRow = 10 To 16
        grdMenu.Cell(flexcpPicture, iRow, 0) = imgList.ListImages("button").Picture
    Next
        
    For iRow = 18 To 21
        grdMenu.Cell(flexcpPicture, iRow, 0) = imgList.ListImages("button").Picture
        ' Print
    Next
    
    grdMenu.Cell(flexcpPictureAlignment, 0, 0, grdMenu.Rows - 1, 0) = flexAlignCenterCenter


End Sub

Public Sub SetupColours()


'    fraAudit.BackColor = sDarkGrey
    
    Select Case grdMenu.RowData(grdMenu.Row)
    ' Select which row is clicked
    
        Case 1  ' Stock Dates
'         fraAudit.BackColor = sDarkGrey
        
        Case 4, "A", "E"  ' Stock Count Entry, Closing Stock
         grdCount.BackColorFixed = sBlue
         grdCount.BackColor = vbWhite
         grdCount.BackColorAlternate = sLightBlue
         grdCount.BackColorSel = sBlue
         
        Case 5, "B"  ' Enter Deliveries, Purchases
         grdCount.BackColorFixed = sGreen
         grdCount.BackColor = vbWhite
         grdCount.BackColorAlternate = sLightGreen
         grdCount.BackColorSel = sGreen
        
        
        Case 6, "D"  ' Enter Till Sales, Till Reconciliation
         grdCount.BackColorFixed = sOrange
         grdCount.BackColor = vbWhite
         grdCount.BackColorAlternate = sLightOrange
         grdCount.BackColorSel = sOrange

        Case "C", "F", "G"  ' Enter Cash, Group Totals, Profit Discrepancy, Summary
'         fraPrint.BackColor = sLightGrey
    
         grdCount.BackColorFixed = sDarkPurple
         grdCount.BackColorSel = sBlue
    
    
    End Select
    

End Sub

Public Function ShowTillSales()
Dim rs As Recordset
Dim iCnt As Integer
Dim iLastPLU As Integer
Dim iLastRow As Integer

    On Error GoTo ErrorHandler
    
    grdCount.Rows = 1
    grdCount.Cols = 0
    
    ' Use Count ID to get Till Sales so far
    
    'ver530
        grdCount.FontSize = 10
    
        SetupCountField frmStockWatch, "No", ""
        SetupCountField frmStockWatch, "Group", ""
        SetupCountField frmStockWatch, "Description", ""
        SetupCountField frmStockWatch, "Sell Price", ""
        SetupCountField frmStockWatch, "Qty", ""
        SetupCountField frmStockWatch, "Gls Price", ""
        SetupCountField frmStockWatch, "Gls Qty", ""
      
        SetupCountField frmStockWatch, "Sell Price 2", ""
        SetupCountField frmStockWatch, "Qty 2", ""
        SetupCountField frmStockWatch, "Gls Price 2", ""
        SetupCountField frmStockWatch, "Gls Qty 2", ""
        
        frmStockWatch.grdCount.ColHidden(frmStockWatch.grdCount.ColIndex("SellPrice2")) = Not bDualPrice
        frmStockWatch.grdCount.ColHidden(frmStockWatch.grdCount.ColIndex("Qty2")) = Not bDualPrice
        frmStockWatch.grdCount.ColHidden(frmStockWatch.grdCount.ColIndex("GlsPrice2")) = Not bDualPrice
        frmStockWatch.grdCount.ColHidden(frmStockWatch.grdCount.ColIndex("GlsQty2")) = Not bDualPrice
        
        
        grdCount.ScrollBars = flexScrollBarVertical
                
        'ver522
        Set rs = SWdb.OpenRecordset("SELECT * FROM (tblClientProductPLUs INNER JOIN tblPLUGroup ON tblClientProductPLUs.PLUGroupID = tblPLUGroup.ID) INNER JOIN tblPLUs ON tblClientProductPLUs.PLUID = tblPLUs.ID WHERE tblClientProductPLUs.ClientID = " & Trim$(lSelClientID) & " AND tblClientPRoductPLUs.Active = true ORDER BY PLUNumber", dbOpenSnapshot)
        ' SQl return list of plus for this clinet sorted by plu#
        If Not (rs.EOF And rs.BOF) Then
    
            rs.MoveFirst
            Do
                
                If iLastPLU <> rs("PLUNumber") Then

                  '  If iLastRow > 0 Then
                  '      If iCnt > 1 Then
                  '          grdCount.Cell(flexcpText, iLastRow, grdCount.ColIndex("description")) = grdCount.Cell(flexcpTextDisplay, iLastRow, grdCount.ColIndex("description")) & "  (" & Trim$(iCnt) & ")"
                  '      End If
                  '  End If
                    
                    If Not IsNull(rs("Glass")) Then
                        iGlass = rs("Glass")
                    Else
                        iGlass = 0
                    End If
                    
'                    If iGlass > 0 Then  '                                                                                                                                                                                                                                                                                                                   rs("SalesQtyDP") - (rs("GlassQtyDP") / iGlass)
'                        grdCount.AddItem rs("PLUNumber") & vbTab & rs("tblPLUGroup.txtDescription") & vbTab & rs("tblPLUs.txtDescription") & vbTab & Format(rs("SellPrice"), "0.00") & vbTab & rs("SalesQty") - (rs("GlassQty") / iGlass) & vbTab & Format(rs("GlassPrice"), "0.00") & vbTab & rs("GlassQty") & vbTab & Format(rs("SellPriceDP"), "0.00") & vbTab & rs("SalesQtyDP") - (rs("GlassQtyDP") / iGlass) & vbTab & Format(rs("GlassPriceDP"), "0.00") & vbTab & rs("GlassQtyDP")
'                    Else
                        grdCount.AddItem rs("PLUNumber") & vbTab & rs("tblPLUGroup.txtDescription") & vbTab & rs("tblPLUs.txtDescription") & vbTab & Format(rs("SellPrice"), "0.00") & vbTab & rs("SalesQty") & vbTab & Format(rs("GlassPrice"), "0.00") & vbTab & rs("GlassQty") & vbTab & Format(rs("SellPriceDP"), "0.00") & vbTab & rs("SalesQtyDP") & vbTab & Format(rs("GlassPriceDP"), "0.00") & vbTab & rs("GlassQtyDP")
'                    End If
                    
                    grdCount.RowData(grdCount.Rows - 1) = rs("tblClientProductPLUs.ID") + 0
                    
                    If Not IsNull(rs("SalesQty")) Then
                        If rs("SalesQty") <> 0 Then
                            grdCount.Cell(flexcpForeColor, grdCount.Rows - 1, 0, grdCount.Rows - 1, grdCount.Cols - 1) = sBlack
                        End If
                    End If
                    
                    If Not IsNull(rs("SalesQtyDP")) Then
                        If rs("SalesQtyDP") <> 0 Then
                            grdCount.Cell(flexcpForeColor, grdCount.Rows - 1, 0, grdCount.Rows - 1, grdCount.Cols - 1) = sBlack
                        End If
                    End If
                    
                    If Not IsNull(rs("GlassQty")) Then
                        If rs("GlassQty") <> 0 Then
                            grdCount.Cell(flexcpForeColor, grdCount.Rows - 1, 0, grdCount.Rows - 1, grdCount.Cols - 1) = sBlack
                        End If
                    End If
                    
                    If Not IsNull(rs("GlassQtyDP")) Then
                        If rs("GlassQtyDP") <> 0 Then
                            grdCount.Cell(flexcpForeColor, grdCount.Rows - 1, 0, grdCount.Rows - 1, grdCount.Cols - 1) = sBlack
                        End If
                    End If
                    
                    iLastPLU = rs("PLUNumber")
                
                End If
                
                rs.MoveNext
            Loop While Not rs.EOF
        
        End If
        
    grdCount.AutoSize 0, grdCount.Cols - 1
    
    
    SetColWidths Me, "grdCount", "Description", False
    
    labelTillCount.Caption = SetCount(grdCount, "No")

    
    ShowTillSales = True
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("ShowTillSales") Then Resume 0
    Resume CleanExit

End Function

Public Function SaveStockCount(lStkId As Long)
Dim rs As Recordset
Dim rsAdd As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblClientProductPLUs")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lStkId
    If Not rs.NoMatch Then
        
        rs.Edit
    
        If bMultipleBars Then
        ' Ver 440
        ' This adds up the totals in all the bars for a particular product
        ' and saves them in the usual spot in table tblClientProductPLUs
            
            Set rsAdd = SWdb.OpenRecordset("SELECT Sum(BarFullQty) as BarFullTot, sum(BarOpen) as BarOpenTot, Sum(BarWeight) as BarWeightTot FROM tblBarCount WHERE ClientProdPLUID = " & Trim$(lStkId), dbOpenSnapshot)
            If Not (rsAdd.EOF And rsAdd.BOF) Then
            
                rs("FullQty") = rsAdd("BarFullTot")
                rs("Open") = rsAdd("BarOpenTot")
                rs("Weight") = rsAdd("BarWeightTot")
                rs.Update
            End If
            
        Else
        ' This just stores the values as before
        
            rs("FullQty") = Val(txtFullQty.Text)
            
            If lblFullQty.Tag = "True" Then
                rs("Open") = Val(txtOpen.Text)
                rs("Weight") = Val(txtWeight.Text)
            End If
            rs.Update
        End If
        
    End If
    
    SaveStockCount = True
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("SaveStockCount") Then Resume 0
    Resume CleanExit

End Function

Public Function GetNextTillNo(lID As Long)
Dim iRow As Integer
    
    If CountInProgress(lSelClientID) Then
    
        
        lID = 0
        
' Ver 2.3
        If grdCount.Row < 1 Then
'============
        ' search grid for next blank Sales Quantity and
        ' force that for an edit
        
            For iRow = 1 To grdCount.Rows - 1
                If grdCount.Cell(flexcpTextDisplay, iRow, grdCount.ColIndex("Qty")) = "" Then
                    lID = grdCount.RowData(iRow)
                    gbOk = GetPLUNo(lID)
                    
                    grdCount.Row = iRow
                            
                    If iRow < grdCount.Rows - 1 Then
                        If Not grdCount.RowIsVisible(iRow + 1) Then
                            grdCount.TopRow = iRow - 1
                        End If
                    End If
                    
                    GetNextTillNo = True
        
                    Exit For
                End If
                
            Next
        
        Else
            
'Ver 2.3
' this is to allow edit of items entered and system will just bump to next item
' in list regardless of having a qty entered or not.
' - Useful for doing multiple changes

            If grdCount.Row < grdCount.Rows - 1 Then
                
                If grdCount.RowData(grdCount.Row + 1) <> 0 Then
                    grdCount.Row = grdCount.Row + 1
                    lID = grdCount.RowData(grdCount.Row)
                    gbOk = GetPLUNo(lID)
        
                    If grdCount.Row < grdCount.Rows - 1 Then
                        If Not grdCount.RowIsVisible(grdCount.Row + 1) Then
                            grdCount.TopRow = grdCount.Row - 1
                        End If
                    End If
                    
                    GetNextTillNo = True
                    
                End If
            
            End If
            
        End If
        
    End If
'========================

End Function

Public Function GetPLUNo(lID As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    'ver530 glass
    'DP
    Set rs = SWdb.OpenRecordset("SELECT tblClientProductPLUs.ID, Glass, GlassPrice, GlassPriceDP, GlassQty, GlassQtyDP, PLUNumber, SellPrice, SellPriceDP, SalesQty, SalesQtyDP, tblPLUs.txtDescription, tblPLUGroup.txtDescription FROM (tblClientProductPLUs INNER JOIN tblPLUGroup ON tblClientProductPLUs.PLUGroupID = tblPLUGroup.ID) INNER JOIN tblPLUs ON tblClientProductPLUs.PLUID = tblPLUs.ID WHERE tblClientProductPLUs.ID = " & Trim$(lID), dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        labelTillDescription.Caption = rs("PLUNumber") & "    " & rs("tblPLUGroup.txtDescription") & "    " & rs("tblPLUs.txtDescription") & "    " & Format(rs("SellPrice"), "0.00")
        labelTillDescription.Tag = Trim$(rs("PLUNumber"))
        ' store the PLU Number here - makes it easier to pick it off
        ' for the save button routine
        
        If Not IsNull(rs("Glass")) Then
            iGlass = rs("Glass")
            
            lblGlass.Visible = True
            txtGlass.Visible = True
            labelGlass1Price.Visible = True
            
   '         lblGlassDP.Visible = bDualPrice
   '         txtGlassDP.Visible = bDualPrice
   '         labelGlassDPPrice.Visible = bDualPrice

        Else
            iGlass = 0
        
            lblGlass.Visible = False
            txtGlass.Visible = False
            labelGlass1Price.Visible = False
            
   '         lblGlassDP.Visible = bDualPrice
   '         txtGlassDP.Visible = bDualPrice
   '         labelGlassDPPrice.Visible = bDualPrice
        
        
        End If
        
        SetMeasureBoxes iGlass
        
        txtSales.Text = ""
        txtSalesDP.Text = ""
        txtGlass.Text = ""
        txtGlassDP.Text = ""
        labelSales1Price.Caption = ""
        labelSalesDPPrice.Caption = ""
        
        
        If rs("SellPrice") > 0 Then
            labelSales1Price.Caption = "@ " & Format(rs("SellPrice"), "Currency")
        End If
        
        If rs("GlassPrice") > 0 Then
            labelGlass1Price.Caption = "@ " & Format(rs("GlassPrice"), "Currency")
        Else
' Ver 547  reenabled these commented out lines
            labelGlass1Price.Caption = ""
            lblGlass.Visible = False
            txtGlass.Visible = False
            labelGlass1Price.Visible = False
        
        End If
        
        If rs("SellPriceDP") > 0 Then
            labelSalesDPPrice.Caption = "@ " & Format(rs("SellPriceDP"), "Currency")
'        Else
'            lblSalesDP.Visible = False
'            txtSalesDP.Visible = False
'            labelSalesDPPrice.Visible = False
        End If
        
        If rs("GlassPriceDP") > 0 Then
            labelGlassDPPrice.Caption = "@ " & Format(rs("GlassPriceDP"), "Currency")
        Else
' Ver 547  reenabled these commented out lines
            labelGlassDPPrice.Caption = ""
            lblGlassDP.Visible = False
            txtGlassDP.Visible = False
            labelGlassDPPrice.Visible = False

        End If
        
        If Not IsNull(rs("SalesQty")) Then
            txtSales.Text = rs("SalesQty")
                
        End If
        
        If Not IsNull(rs("GlassQty")) Then
            txtGlass.Text = rs("GlassQty")
        End If
        
        If Not IsNull(rs("SalesQtyDP")) Then
                If Not IsNull(rs("SalesQtyDP")) Then
                    If rs("SalesQtyDP") > 0 Then
'                        txtSalesDP.Text = rs("SalesQtyDP") - (rs("GlassQtyDP") / iGlass)
'                    Else
                        txtSalesDP.Text = rs("SalesQtyDP")
                    End If
                Else
                    txtSalesDP.Text = rs("SalesQtyDP")
                End If
        End If
        
        If Not IsNull(rs("GlassQtyDP")) Then
            txtGlassDP.Text = rs("GlassQtyDP")
        End If
    
    End If
    
    GetPLUNo = True
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetPLUNo") Then Resume 0
    Resume CleanExit

End Function

Public Function RefreshTillSales(lID As Long, lCLId As Long, sPLU As String)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    ' ver530 glasses included here
    
    
    Set rs = SWdb.OpenRecordset("SELECT * FROM tblClientProductPLUs WHERE ClientID = " & Str$(lCLId) & " AND PLUNumber = " & sPLU & " ORDER BY SalesQty, SalesQtyDP Desc")
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        'rs.MoveFirst
        
'        If iGlass = 0 Then
            grdCount.Cell(flexcpText, grdCount.FindRow(lID), grdCount.ColIndex("Qty")) = rs("SalesQty")
            grdCount.Cell(flexcpText, grdCount.FindRow(lID), grdCount.ColIndex("Qty2")) = rs("SalesQtyDP")
'        Else
'            grdCount.Cell(flexcpText, grdCount.FindRow(lID), grdCount.ColIndex("Qty")) = rs("SalesQty") - (rs("GlassQty") / iGlass)
'            grdCount.Cell(flexcpText, grdCount.FindRow(lID), grdCount.ColIndex("Qty2")) = rs("SalesQtyDP") - (rs("GlassQtyDP") / iGlass)
        
'        End If
        
        grdCount.Cell(flexcpText, grdCount.FindRow(lID), grdCount.ColIndex("GlsQty")) = rs("GlassQty")
        grdCount.Cell(flexcpText, grdCount.FindRow(lID), grdCount.ColIndex("GlsQty2")) = rs("GlassQtyDP")
        
        grdCount.Cell(flexcpForeColor, grdCount.FindRow(lID), 0, grdCount.FindRow(lID), grdCount.Cols - 1) = sBlack
    End If
    
    RefreshTillSales = True
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("RefreshTillSales") Then Resume 0
    Resume CleanExit




End Function
Public Function ShowTotalStockCount()
Dim rs As Recordset
Dim iLastGroup As Integer

    On Error GoTo ErrorHandler
    
    ' Ver440
    ' This is basically the same as originally except its renamed from ShowStockCount
    
    grdCount.Rows = 1
    grdCount.Cols = 0
    
    ' Use Count ID to get Till Sales so far
    
    SetupCountField frmStockWatch, "Code", ""
    SetupCountField frmStockWatch, "Description", ""
    SetupCountField frmStockWatch, "Size", ""
    SetupCountField frmStockWatch, "Full Qty", ""
    SetupCountField frmStockWatch, "Open Items", ""
    SetupCountField frmStockWatch, "Weight", ""

    grdCount.ScrollBars = flexScrollBarVertical
    
'    Set rs = SWdb.OpenRecordset("SELECT * FROM (tblClientProductPLUs INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE tblClientProductPLUs.ClientID = " & Trim$(lSelClientID) & " AND tblProducts.chkActive = true ORDER BY tblProducts.cboGroups, tblProducts.txtDescription, txtSize", dbOpenSnapshot)
'ver520
    Set rs = SWdb.OpenRecordset("SELECT * FROM (tblClientProductPLUs INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE tblClientProductPLUs.ClientID = " & Trim$(lSelClientID) & " AND tblProducts.chkActive = true AND Active = true ORDER BY tblProducts.cboGroups, tblProducts.txtDescription, txtSize", dbOpenSnapshot)
    ' SQl return list of plus for this clinet sorted by plu#
    
    If Not (rs.EOF And rs.BOF) Then

        rs.MoveFirst
        Do
            
            If iLastGroup <> rs("cboGroups") Then

                grdCount.AddItem vbTab & rs("cboGroups") & "  " & rs("tblProductGroup.txtDescription")
                iLastGroup = rs("cboGroups")
                grdCount.Cell(flexcpAlignment, grdCount.Rows - 1, 1) = flexAlignLeftCenter
                grdCount.Cell(flexcpBackColor, grdCount.Rows - 1, 0, grdCount.Rows - 1, grdCount.Cols - 1) = &HC0C0C0
                grdCount.Cell(flexcpForeColor, grdCount.Rows - 1, 0, grdCount.Rows - 1, grdCount.Cols - 1) = sBlack

            End If
            
            
            If IsNull(rs("FullQty")) Then
                grdCount.AddItem rs("txtCode") & vbTab & rs("tblProducts.txtDescription") & vbTab & rs("txtSize")
                grdCount.Cell(flexcpForeColor, grdCount.Rows - 1, 0, grdCount.Rows - 1, grdCount.Cols - 1) = sDarkGrey
            
            ElseIf rs("FullQty") = Empty And rs("Open") = Empty Then
                grdCount.AddItem rs("txtCode") & vbTab & rs("tblProducts.txtDescription") & vbTab & rs("txtSize") & vbTab & "0" & vbTab & "0" & vbTab & "0"
                grdCount.Cell(flexcpForeColor, grdCount.Rows - 1, 0, grdCount.Rows - 1, grdCount.Cols - 1) = sBlack
            
                
            Else
                grdCount.AddItem rs("txtCode") & vbTab & rs("tblProducts.txtDescription") & vbTab & rs("txtSize") & vbTab & rs("FullQty") & vbTab & rs("Open") & vbTab & rs("Weight")
                grdCount.Cell(flexcpForeColor, grdCount.Rows - 1, 0, grdCount.Rows - 1, grdCount.Cols - 1) = sBlack
            
            End If
            
            grdCount.Cell(flexcpData, grdCount.Rows - 1, grdCount.ColIndex("description")) = rs("ProductID") + 0
            ' Ver 440
            
            grdCount.RowData(grdCount.Rows - 1) = rs("tblClientProductPLUs.ID") + 0
                
            rs.MoveNext
        Loop While Not rs.EOF
    
    End If
        
    SetColWidths Me, "grdCount", "Description", False
    
    labelStockCount.Caption = SetCount(grdCount, "FullQty")
    
    
    ShowTotalStockCount = True
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("ShowTotalStockCount") Then Resume 0
    Resume CleanExit

End Function

Public Function GetNextStockItem(lID As Long)
Dim iRow As Integer
Dim iFromThisRowForward As Integer

    If CountInProgress(lSelClientID) Then
    
      If grdCount.Row > -1 Then
        
        iFromThisRowForward = grdCount.Row
        ' set pointer so we look forward from this point
        
        lID = 0
        
        ' search grid for next blank Sales Quantity and
        ' force that for an edit
    
LookForNextItemThatsBlank:

        For iRow = iFromThisRowForward To grdCount.Rows - 1
            If grdCount.Cell(flexcpTextDisplay, iRow, grdCount.ColIndex("FullQty")) = "" Then
            ' check is full quantity blank
            
                If grdCount.RowData(iRow) <> Empty Then
                ' but now also check if its a group header row and ignore if it is
                
                    lID = grdCount.RowData(iRow)
                    ' ok save the id here...
                    
                    If bMultipleBars Then
                        gbOk = GetStockItem(lID)
                    
'                        gbOk = GetBarStockItem(lID)
                        ' now get the stock details of that product
                    
                    Else
                        gbOk = GetStockItem(lID)
                        ' now get the stock details of that product
                    
                    End If
                    
                    grdCount.Row = iRow
                    ' and point to the row itself
                    
                    If iRow < grdCount.Rows - 1 Then
                        If Not grdCount.RowIsVisible(iRow + 1) Then
                            grdCount.TopRow = iRow - 1
                        End If
                    End If
                    ' make sure the row is visible in case its off the page
                    
                    GetNextStockItem = True
                    
                    'ver440
'                    Exit Function
                    Exit For
                End If
                
            End If
            
        Next
 
 
 
 ' TESTING ver 5.2
'        If iFromThisRowForward <> 0 Then
'            iFromThisRowForward = 0
'            GoTo LookForNextItemThatsBlank
'
'        End If
        
        GetNextStockItem = True
    
      
      End If
      
    End If
    
End Function

Public Function GetStockItem(lID As Long)
Dim rs As Recordset
Dim rsBar As Recordset

    On Error GoTo ErrorHandler
    
    If lID > 0 Then
    
        Set rs = SWdb.OpenRecordset("SELECT txtFullWeight, txtEmptyWeight, chkOpenItem, txtCode, tblClientProductPLUs.ID, tblProducts.txtDescription, tblProductGroup.txtDescription, tblProducts.txtsize, FullQty, Open, Weight FROM (tblClientProductPLUs INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE tblClientProductPLUs.ID = " & Trim$(lID), dbOpenSnapshot)
    
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
            labelDescription.Caption = rs("txtCode") & "    " & rs("tblProductGroup.txtDescription") & "    " & rs("tblProducts.txtDescription") & "    " & rs("txtSize")
                    
            
            lblFullQty.Tag = rs("chkOpenItem")
            txtOpen.Visible = rs("chkOpenItem")
            txtWeight.Visible = rs("chkOpenItem")
            lblOpen.Visible = rs("chkOpenItem")
            lblWeight.Visible = rs("chkOpenItem")
            
            labelFullWeight.Caption = rs("txtFullWeight")
            labelEmptyWeight.Caption = rs("txtEmptyWeight")
            lblFullWeight.Visible = rs("chkOpenItem")
            lblEmptyWeight.Visible = rs("chkOpenItem")
            labelFullWeight.Visible = rs("chkOpenItem")
            labelEmptyWeight.Visible = rs("chkOpenItem")
            
                
            If lBarID > 0 Then
            
                Set rsBar = SWdb.OpenRecordset("SELECT * FROM tblBarCount WHERE ClientProdPLUID = " & Trim$(rs("ID")) & " AND BarID = " & Trim$(lBarID), dbOpenSnapshot)
                If Not (rsBar.EOF And rsBar.BOF) Then
                
                    rsBar.MoveFirst
                    
                    If Not IsNull(rsBar("BarFullQty")) Then
                        txtFullQty.Text = rsBar("BarFullQty")
'                        dblFull = rsBar("BarFullQty")
                    Else
                        txtFullQty.Text = ""
'                        dblFull = 0
                    End If
                
                    If Not IsNull(rsBar("BarOpen")) Then
                        txtOpen.Text = rsBar("BarOpen")
'                        dblOpen = rsBar("BarOpen")
                    Else
                        txtOpen.Text = ""
'                        dblOpen = 0
                    End If
                
                    If Not IsNull(rsBar("BarWeight")) Then
                        txtWeight.Text = rsBar("BarWeight")
'                        dblWeight = rsBar("BarWeight")
                    Else
                        txtWeight.Text = ""
'                        dblWeight = 0
                    End If
                End If
                
                    
            Else
            
                If Not IsNull(rs("FullQty")) Then
                    txtFullQty.Text = rs("FullQty")
'                    dblFull = rs("FullQty")
                Else
                    txtFullQty.Text = ""
'                    dblFull = 0
                End If
            
                If Not IsNull(rs("Open")) Then
                    txtOpen.Text = rs("Open")
'                    dblOpen = rs("Open")
                Else
                    txtOpen.Text = ""
'                    dblOpen = 0
                End If
            
                If Not IsNull(rs("Weight")) Then
                    txtWeight.Text = rs("Weight")
'                    dblweight = rs("Weight")
                Else
                    txtWeight.Text = ""
'                    dblweight = 0
                End If
            
            End If
            
            lStkId = rs("ID") + 0
        
        End If
        rs.Close
    
    End If
    
    GetStockItem = True

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetStockItem") Then Resume 0
    Resume CleanExit

End Function

Public Function UpdateStockCount(lID As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    If Not bMultipleBars Then
    
        Set rs = SWdb.OpenRecordset("tblClientProductPLUs")
        rs.Index = "PrimaryKey"
        rs.Seek "=", lID
        If Not rs.NoMatch Then
            'rs.MoveFirst
            grdCount.Cell(flexcpText, grdCount.FindRow(lID), grdCount.ColIndex("FullQty")) = rs("FullQty")
            grdCount.Cell(flexcpText, grdCount.FindRow(lID), grdCount.ColIndex("OpenItems")) = rs("Open")
            grdCount.Cell(flexcpText, grdCount.FindRow(lID), grdCount.ColIndex("Weight")) = rs("Weight")
            grdCount.Cell(flexcpForeColor, grdCount.FindRow(lID), 0, grdCount.FindRow(lID), grdCount.Cols - 1) = sBlack
        End If
    
    Else
    
        Set rs = SWdb.OpenRecordset("SELECT * FROM tblBarCount WHERE ClientProdPLUID = " & Trim$(lID) & " AND BarID = " & Trim$(lBarID), dbOpenSnapshot)
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
            
            grdCount.Cell(flexcpText, grdCount.FindRow(lID), grdCount.ColIndex("FullQty")) = rs("BarFullQty")
            grdCount.Cell(flexcpText, grdCount.FindRow(lID), grdCount.ColIndex("OpenItems")) = rs("BarOpen")
            grdCount.Cell(flexcpText, grdCount.FindRow(lID), grdCount.ColIndex("Weight")) = rs("BarWeight")
            grdCount.Cell(flexcpForeColor, grdCount.FindRow(lID), 0, grdCount.FindRow(lID), grdCount.Cols - 1) = sBlack
        End If
    
    
    End If
    
    UpdateStockCount = True
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("UpdateStockCount") Then Resume 0
    Resume CleanExit

End Function


Private Sub txtVoucherSales_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtVoucherSales_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            KeyAscii = 0
        ' kill beep

    ElseIf KeyAscii <> 46 Then
        
        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
        
    ElseIf InStr(txtVoucherSales, ".") > 0 Then
            
        KeyAscii = 0
    End If

End Sub

Private Sub txtVoucherSales_KeyUp(KeyCode As Integer, Shift As Integer)
    gbOk = getProjectedSales()

End Sub

Private Sub txtVoucherSales_LostFocus()
    lblVoucherSales.ForeColor = sBlack

End Sub

Private Sub txtWastage_GotFocus()
    gbOk = bSetupControl(Me)

End Sub
Private Sub txtSurplus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            KeyAscii = 0
        ' kill beep
        
        bSetFocus Me, "cmdCashSave"

    ElseIf KeyAscii <> 46 Then
        
        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
        
    ElseIf InStr(txtSurplus, ".") > 0 Then
            
        KeyAscii = 0
    End If

End Sub
Private Sub txtSurplus_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtWastage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            KeyAscii = 0
        ' kill beep

    ElseIf KeyAscii <> 46 Then
        
        KeyAscii = CharOk(KeyAscii, 0, ".") ' 0 = no only, 1 = char only, 2 = both
        
    ElseIf InStr(txtWastage, ".") > 0 Then
            
        KeyAscii = 0
    End If

End Sub

Private Sub txtWastage_KeyUp(KeyCode As Integer, Shift As Integer)
    gbOk = getProjectedSales()

End Sub

Private Sub txtSurplus_KeyUp(KeyCode As Integer, Shift As Integer)
    gbOk = getProjectedSales()

End Sub

Private Sub txtWastage_LostFocus()
    lblWastage.ForeColor = sBlack

End Sub

Private Sub txtSurplus_LostFocus()
    txtSurpluslbl.ForeColor = sBlack
    
    
End Sub

Private Sub txtWeight_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtWeight_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDown Then
        bSetFocus Me, "grdCount"
    ElseIf KeyCode = vbKeyUp Then
        bSetFocus Me, "txtOpen"
    End If
    

End Sub

Private Sub txtWeight_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        
        KeyAscii = 0
        ' kill beep
        
        If bAddIt Then
        
            bAddIt = False
            txtWeight = AddUp(Trim$(txtWeight.Text))
            txtWeight.Width = iBoxWidth
            txtWeight.ForeColor = sBlack
        
        End If
            
        If CombWeightOK(Val(txtOpen), Val(txtWeight)) Then
            labelMsg.Caption = ""
            cmdSaveStock_Click
        End If
       
    ElseIf KeyAscii = 43 Then
    ' +
        txtWeight.SelStart = Len(txtWeight)
        
        txtWeight.Width = SetBoxWidth(txtWeight)
        
        txtWeight.ForeColor = vbBlue
        bAddIt = True
        
    Else
        KeyAscii = CharOk(KeyAscii, 0, "") ' 0 = no only, 1 = char only, 2 = both

        txtWeight.Width = SetBoxWidth(txtWeight)
    
        cmdSaveStock.Enabled = True
    
    End If

End Sub

Public Function ShowDeliveries()
Dim rs As Recordset
Dim iLastGroup As Integer

    On Error GoTo ErrorHandler
    
    grdCount.Rows = 1
    grdCount.Cols = 0
    
    ' Use Count ID to get Till Sales so far
    
    SetupCountField frmStockWatch, "Ref", ""
    SetupCountField frmStockWatch, "Code", ""
    SetupCountField frmStockWatch, "Description", ""
    SetupCountField frmStockWatch, "Size", ""
    SetupCountField frmStockWatch, "Delivery Note", ""
    SetupCountField frmStockWatch, "Quantity", ""
    SetupCountField frmStockWatch, "Cost", ""
    SetupCountField frmStockWatch, "Free", ""
    grdCount.ScrollBars = flexScrollBarVertical
    
                    
'Ver522
'    Set rs = SWdb.OpenRecordset("SELECT tblClientPRoductPLUs.ID, Ref, txtCode, txtDescription, txtSize, tblDeliveries.DeliveryNote, Quantity, PurchasePrice, Free, tblDeliveries.ID FROM (tblClientProductPLUs LEFT JOIN tblDeliveries ON tblClientProductPLUs.ID = tblDeliveries.ClientProdPLUID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID WHERE tblClientProductPLUs.ClientID = " & Trim$(lSelClientID) & " ORDER BY tblProducts.cboGroups, tblProducts.txtDescription, txtSize, ref", dbOpenSnapshot)

    Set rs = SWdb.OpenRecordset("SELECT tblClientPRoductPLUs.ID, Ref, txtCode, txtDescription, txtSize, tblDeliveries.DeliveryNote, Quantity, PurchasePrice, Free, tblDeliveries.ID FROM (tblClientProductPLUs LEFT JOIN tblDeliveries ON tblClientProductPLUs.ID = tblDeliveries.ClientProdPLUID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID WHERE tblClientProductPLUs.ClientID = " & Trim$(lSelClientID) & " AND tblClientPRoductPLUs.Active = true ORDER BY tblProducts.cboGroups, tblProducts.txtDescription, txtSize, ref", dbOpenSnapshot)
'------

    ' SQl return list of plus for this clinet sorted by plu#
    If Not (rs.EOF And rs.BOF) Then

        rs.MoveFirst
        Do
            
            
'Ver 2.1.0
'            grdCount.AddItem rs("Ref") & vbTab & rs("txtCode") & vbTab & rs("txtDescription") & vbTab & rs("txtSize") & vbTab & rs("DeliveryNote") & vbTab & rs("Quantity") & vbTab & rs("Cost") & vbTab & rs("Free")
            
            
'   Ver 2.2.0
'            grdCount.AddItem rs("Ref") & vbTab & rs("txtCode") & vbTab & rs("txtDescription") & vbTab & rs("txtSize") & vbTab & rs("DeliveryNote") & vbTab & rs("Quantity") & vbTab & Format(rs("Cost"), "0.00") & vbTab & rs("Free")
'   =========
            grdCount.AddItem rs("Ref") & vbTab & rs("txtCode") & vbTab & rs("txtDescription") & vbTab & rs("txtSize") & vbTab & rs("DeliveryNote") & vbTab & rs("Quantity") & vbTab & Format(rs("PurchasePrice"), "0.00") & vbTab & rs("Free")
'=======

            grdCount.Cell(flexcpForeColor, grdCount.Rows - 1, 0, grdCount.Rows - 1, grdCount.Cols - 1) = sBlack
            grdCount.Cell(flexcpData, grdCount.Rows - 1, 0) = rs("tblDeliveries.ID") + 0
                
            grdCount.RowData(grdCount.Rows - 1) = rs("tblClientProductPLUs.ID") + 0
            
            rs.MoveNext
        Loop While Not rs.EOF
    
    End If
        
    grdCount.AutoSize 0, grdCount.Cols - 1
    
        
    SetColWidths Me, "grdCount", "Description", True
    
    labelDelCount.Caption = SetCount(grdCount, "Quantity")
    
    
    ShowDeliveries = True
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("ShowDeliveries") Then Resume 0
    Resume CleanExit

End Function

Public Function GetNextDelivery(lID As Long)
Dim iRow As Integer
    
    If CountInProgress(lSelClientID) Then
    
        lID = 0
        
        ' search grid for next blank Sales Quantity and
        ' force that for an edit
    
            For iRow = 1 To grdCount.Rows - 1
                If grdCount.Cell(flexcpTextDisplay, iRow, grdCount.ColIndex("Quantity")) = "" Then
                        
                    gbOk = InitDeliveryItem()
                    
                    lID = grdCount.RowData(iRow)
                    
                    gbOk = GetDeliveryItem(False, lID)
                    
                    grdCount.Row = iRow
                            
                    If iRow < grdCount.Rows - 1 Then
                        If Not grdCount.RowIsVisible(iRow + 1) Then
                            grdCount.TopRow = iRow - 1
                        End If
                    End If
                    GetNextDelivery = True
                                    
                    bSetFocus Me, "txtQty"
                    
                    Exit Function
                
                End If
                
            Next
    
    End If
    
    
    
    
End Function

Public Function GetDeliveryItem(bWhich As Boolean, lID As Long)
'Public Function GetDeliveryItem(bWhich As Boolean, lID As Long, lCode As Long)
Dim rs As Recordset

    
    On Error GoTo ErrorHandler
    
    If bWhich Then
    ' edit of a delivery
    
'Ver 2.2.0 Change cost to purchaseprice

'        Set rs = SWdb.OpenRecordset("SELECT ref, txtCode, txtSize, Quantity, DeliveryNote, Cost, Free, tblClientProductPLUs.ID, tblProducts.txtDescription, tblProductGroup.txtDescription FROM ((tblClientProductPLUs INNER JOIN tblDeliveries ON tblClientProductPLUs.ID = tblDeliveries.ClientProdPLUID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE tblClientProductPLUs.ClientID = " & Trim$(lSelClientID) & " AND tblDeliveries.ID = " & Trim$(lID), dbOpenSnapshot)
        Set rs = SWdb.OpenRecordset("SELECT ref, txtCode, txtSize, Quantity, DeliveryNote, PurchasePrice, Free, tblClientProductPLUs.ID, tblProducts.txtDescription, tblProductGroup.txtDescription FROM ((tblClientProductPLUs INNER JOIN tblDeliveries ON tblClientProductPLUs.ID = tblDeliveries.ClientProdPLUID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE tblClientProductPLUs.ClientID = " & Trim$(lSelClientID) & " AND tblDeliveries.ID = " & Trim$(lID), dbOpenSnapshot)
    
    Else
    ' new item
        
        Set rs = SWdb.OpenRecordset("SELECT ref, PurchasePrice, txtCode, txtSize, Quantity, DeliveryNote, Free,tblClientProductPLUs.ID, tblProducts.txtDescription, tblProductGroup.txtDescription FROM ((tblClientProductPLUs LEFT JOIN tblDeliveries ON tblClientProductPLUs.ID = tblDeliveries.ClientProdPLUID) INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE tblClientProductPLUs.ID = " & Trim$(lID), dbOpenSnapshot)
    End If
    
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        labelDelItem.Caption = rs("txtCode") & "    " & rs("tblProductGroup.txtDescription") & "    " & rs("tblProducts.txtDescription") & "    " & rs("txtSize")
        
        If Not IsNull(rs("Quantity")) Then
            txtQty.Text = rs("Quantity")
            If Not IsNull(rs("Free")) Then
                txtFree.Text = rs("Free")
            End If

'Ver 2.2.0

            If Not IsNull(rs("PurchasePrice")) Then
                txtDelCost.Text = Format(rs("PurchasePrice"), "0.00")
'                curDefaultCost = Format(rs("PurchasePrice"), "0.00")
            End If
            
'            txtDelCost.Text = Format(rs("Cost"), "0.00")
'            curDefaultCost = Format(rs("Cost"), "0.00")
        Else
            If Not IsNull(rs("PurchasePrice")) Then
                txtDelCost.Text = Format(rs("PurchasePrice"), "0.00")
'                curDefaultCost = Format(rs("PurchasePrice"), "0.00")
'            ElseIf Not IsNull(rs("Cost")) Then
'                txtDelCost.Text = Format(rs("Cost"), "0.00")
'                curDefaultCost = Format(rs("Cost"), "0.00")
            End If
            
            txtQty.Text = ""
        
        End If
    
        If IsNull(rs("ref")) Or rs("ref") = 1 Then
            cboDelivery.ListIndex = 0
            txtDelOther.Visible = False
            txtDelOther.Text = ""
        Else
            cboDelivery.ListIndex = 1
            txtDelOther.Visible = True
            txtDelOther.Text = rs("DeliveryNote") & ""
        
        End If
        
        GetDeliveryItem = True
    
    End If
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetDeliveryItem") Then Resume 0
    Resume CleanExit

End Function

Public Function SaveDelivery(lID As Long, lNextDelID As Long, iDelivery As Integer, iRef As Integer)
Dim rs As Recordset

    
'Ver 2.2.0
' Cost needs to be saved directly to the table tblClientProductPLUs
    
    
    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblDeliveries")
    rs.Index = "PrimaryKey"
    
    'NEW MAIN
    If iDelivery = 0 And iRef = 0 Then
    ' new
        rs.AddNew
        rs("Ref") = 1
        rs("DeliveryNote") = cboDelivery.Text

        
    'EDIT MAIN
    ElseIf iDelivery = 0 And iRef = 1 Then
    ' edit

        rs.Seek "=", lID
        If Not rs.NoMatch Then
        ' Found a match but now make sure the same ref
        
            rs.Edit
        End If
        

    'NEW OTHER
    ElseIf iDelivery = 1 And iRef = 1 Then
    ' New Other

        rs.AddNew
        rs("Ref") = 2
        rs("DeliveryNote") = txtDelOther.Text
    
    'UPDATE / DELETE
    ElseIf iDelivery = 1 And iRef = 2 Then
    ' Update/delete

        If bEditOther Then
            
            rs.Seek "=", lID
            If Not rs.NoMatch Then
            ' Found a match but now make sure the same ref
            
                rs.Edit
                rs("DeliveryNote") = txtDelOther.Text
            
            End If
            
        Else
        ' delete record
        
            rs.Seek "=", lID
            If Not rs.NoMatch Then
            ' Found a match but now make sure the same ref
            
                rs.Delete
                
                LogMsg Me, "Delivery: " & labelDelItem.Caption & " Deleted " & Replace(lblClient.Tag, "_", " "), labelDelItem.Caption & " Qty:" & txtQty & " free:" & txtFree & " Cost:" & txtDelCost

                GoTo CleanExit
                
            End If
        
        End If
    
    End If
    
    rs("ClientProdPLUID") = lNextDelID
    rs("Quantity") = Val(txtQty.Text)
    rs("Free") = Val(txtFree.Text)
    
'    rs("Cost") = Format(Val(txtDelCost.Text), "0.00")
    rs.Update
    rs.Bookmark = rs.LastModified
    lID = rs("ID") + 0


CleanExit:
    SaveDelivery = True
    rs.Close
    
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
        If CheckDBError("SaveDelivery") Then Resume 0
        Resume CleanExit
    
    End If
    

End Function

Public Function UpdateDelivery(lID As Long, lNextID As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblDeliveries")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lID
    If Not rs.NoMatch Then
        'rs.MoveFirst

        grdCount.Cell(flexcpText, grdCount.Row, grdCount.ColIndex("Ref")) = Trim$(rs("Ref"))
        
        grdCount.Cell(flexcpText, grdCount.Row, grdCount.ColIndex("Quantity")) = Trim$(rs("Quantity"))
        grdCount.Cell(flexcpText, grdCount.Row, grdCount.ColIndex("Free")) = Trim$(rs("Free"))
        grdCount.Cell(flexcpText, grdCount.Row, grdCount.ColIndex("DeliveryNote")) = rs("DeliveryNote")
        
'Ver 2.2.0

        grdCount.Cell(flexcpText, grdCount.Row, grdCount.ColIndex("Cost")) = GetPurchasePrice(lNextID)
        grdCount.Cell(flexcpData, grdCount.Row, 0) = lID
        ' Save the new Index ID from the Delivieries table
        
        grdCount.Cell(flexcpForeColor, grdCount.Row, 0, grdCount.Row, grdCount.Cols - 1) = sBlack
    
    End If
    
    SetColWidths Me, "grdCount", "Description", True
    
    
    UpdateDelivery = True
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("UpdateDelivery") Then Resume 0
    Resume CleanExit


End Function

Public Function ShowCashTakings(lID As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblDates")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lID
    If Not rs.NoMatch Then
        txtActual.Text = Format(rs("Actual"), "0.00")
        txtStaff.Text = Format(rs("Staff"), "0.00")
        txtComplimentary.Text = Format(rs("Complimentary"), "0.00")
        txtWastage.Text = Format(rs("Wastage"), "0.00")
        txtOverRings.Text = Format(rs("OverRings"), "0.00")
        txtPromotions.Text = Format(rs("Promotions"), "0.00")
        txtOffLicense.Text = Format(rs("OffLicense"), "0.00")
        txtVoucherSales.Text = Format(rs("VoucherSales"), "0.00")
        txtKitchen.Text = Format(rs("Kitchen"), "0.00")
        If txtOtherLbl.Text <> "Other" Then
            txtOtherLbl.Text = rs("OtherTitle")
        End If
        txtOther.Text = Format(rs("Other"), "0.00")

'ver 547
        If Not IsNull(rs("SurplusTitle")) Then
            If rs("SurplusTitle") <> "Cash €" Then
                If Not IsNull(rs("SurplusTitle")) Then
                    txtSurpluslbl.Text = rs("SurplusTitle")
                End If
            End If
        End If
        txtSurplus.Text = Format(rs("Surplus"), "0.00")
    
    End If
    
    ShowCashTakings = True
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("ShowCashTakings") Then Resume 0
    Resume CleanExit


End Function

Public Function SaveCashTakings(lID As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblDates")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lID
    If Not rs.NoMatch Then
        rs.Edit
        rs("Actual") = Val(txtActual.Text)
        rs("Staff") = Val(txtStaff.Text)
        rs("Complimentary") = Val(txtComplimentary.Text)
        rs("Wastage") = Val(txtWastage.Text)
        rs("OverRings") = Val(txtOverRings.Text)
        rs("Promotions") = Val(txtPromotions.Text)
        rs("OffLicense") = Val(txtOffLicense.Text)
        rs("VoucherSales") = Val(txtVoucherSales.Text)
        rs("Kitchen") = Val(txtKitchen.Text)
        rs("Other") = Val(txtOther.Text)
        rs("OtherTitle") = Trim$(txtOtherLbl.Text)
        rs("Surplus") = Val(txtSurplus.Text)
        rs("SurplusTitle") = Trim$(txtSurpluslbl.Text)
        
        rs.Update
    End If
    
    SaveCashTakings = True
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("SaveCashTakings") Then Resume 0
    Resume CleanExit


End Function


Public Function CombWeightOK(iOpen As Integer, iCombWeight As Long)

    If iCombWeight < (iOpen * Val(labelEmptyWeight)) Then
        labelMsg.Caption = "Combined weight is Too Low"
        bSetFocus Me, "txtOpen"
    ElseIf iCombWeight > (iOpen * Val(labelFullWeight)) Then
        labelMsg.Caption = "Combined weight is Too High            "
        bSetFocus Me, "txtOpen"
    Else
        CombWeightOK = True
    End If

End Function

Public Sub SetUpStockTakeForms()

    fraStock.Top = 0
    fraDates.Top = 0
    fraSales.Top = 0
    fraDelivery.Top = 0

End Sub
Public Function DeliveryFieldsOK()

    If lNextDelID <> 0 Then
        If Val(txtQty) > 1000 Then
            If MsgBox("Are you sure Quantity " & txtQty & " is correct?", vbDefaultButton1 + vbYesNo + vbQuestion, "Warning: Quantity > 1000") = vbNo Then
                bSetFocus Me, "txtQty"
                Exit Function
            End If
        End If
        
        If cboDelivery.ListIndex = 0 Or (cboDelivery.ListIndex = 1 And Trim$(txtDelOther.Text) <> "") Then
            
            If Val(txtDelCost.Text) = 0 And Val(txtQty.Text) > 0 Then
                If MsgBox("Are you sure you want to set the Cost to 0.0?", vbDefaultButton1 + vbYesNo + vbQuestion, "Set Cost to Zero") = vbYes Then
                    DeliveryFieldsOK = True
                End If
            
            Else
                DeliveryFieldsOK = True
            End If
                    
        Else
            MsgBox "Please enter a Delivery Note for 'Other'"
            bSetFocus Me, "txtDelOther"
        
        End If
    Else
        MsgBox "Please Select an Item from the List"
    End If

End Function

Public Function StockFieldsOk()
Dim sMsg As String

    If lStkId > 0 Then
        
        sMsg = "Are you sure the full quantity number of " & txtFullQty & " is correct?"
        
        If lblFullQty.Tag = True Then
            If Val(txtFullQty) > 1000 Then
                If MsgBox(sMsg, vbDefaultButton1 + vbYesNo + vbQuestion, "Warning: Full quantity > 1000") = vbNo Then
                    bSetFocus Me, "txtFullQty"
                    Exit Function
                End If
            End If
        
        ElseIf Val(txtFullQty) > 10000 Then
            If MsgBox(sMsg, vbDefaultButton1 + vbYesNo + vbQuestion, "Warning: Full quantity > 10000") = vbNo Then
                bSetFocus Me, "txtFullQty"
                Exit Function
            End If
        End If
        
        If Val(txtOpen) > 50 Then
            If MsgBox("Are you sure an open quantity of " & txtOpen & " is Correct?", vbDefaultButton1 + vbYesNo + vbQuestion, "Warning: Open quantity > 50") = vbNo Then
                bSetFocus Me, "txtOpen"
                Exit Function
            End If
        End If
        
        
        ' check full qty and warn if > 10000 for single units
        ' or > 100 for kegs
        
        ' check open qty and warn if > 20
        
        
        If CombWeightOK(Val(txtOpen), Val(txtWeight)) Then
            StockFieldsOk = True
        End If
        
    Else
        MsgBox "Please Select an Item from the List, or Add a new one in Client Products/PLUs"
        bSetFocus Me, "txtFullQty"
    End If
End Function

Public Function InitStockEntry()

'    btnBars.Visible = bMultipleBars
 '   txtCode.Text = ""
    labelDescription.Caption = ""
    txtFullQty.Text = ""
    txtOpen.Text = ""
    txtWeight.Text = ""
    labelFullWeight.Caption = ""
    labelEmptyWeight.Caption = ""
    labelMsg.Caption = ""
    cmdSaveStock.Enabled = False
    lStkId = 0
    
End Function

Private Sub txtWeight_KeyUp(KeyCode As Integer, Shift As Integer)
    txtWeight.Width = SetBoxWidth(txtWeight)

End Sub

Private Sub txtWeight_LostFocus()
    
    bAddIt = False
    lblWeight.ForeColor = sBlack

End Sub

Public Function SetCount(grd As VSFlexGrid, sGridColKey As String) As String
Dim iRow As Integer
Dim iCnt As Integer
Dim itotal As Integer

    With grd
    
    For iRow = 1 To .Rows - 1
    
        If .RowData(iRow) <> Empty Then
            itotal = itotal + 1
            If .Cell(flexcpText, iRow, .ColIndex(sGridColKey)) <> "" Then
                iCnt = iCnt + 1
            End If
        End If
        
    Next

    SetCount = "  " & Trim$(iCnt) & " of " & Trim$(itotal) & " Complete  "

    End With

End Function

Public Function InitDeliveryItem()

'    cmdAddNew.Visible = False
    
    SetDeliveryLocked True
    txtDelOther.Visible = False
    txtDelOther.Text = ""
    labelDelItem.Caption = ""
    cboDelivery.ListIndex = 0
    txtQty.Text = ""
    txtDelCost.Text = ""
    txtFree.Text = ""
'    cmdAddNew.Enabled = False
    cmdDelSave.Enabled = False
    lDelID = 0
    lNextDelID = 0
    

' ver 2.3
'    cmdAddNew.Caption = "&Add New"
    
End Function

Public Function TillFieldsOK()

    If lTillID <> 0 Then
        
        If Val(txtSales) > 20000 Or Val(txtSalesDP) > 20000 Or Val(txtGlass) > 20000 Or Val(txtGlassDP) > 20000 Then
            If MsgBox("Are you sure the Sales quantity of " & txtSales & " is correct?", vbDefaultButton1 + vbYesNo + vbQuestion, "Warning: Sales Quantity > 20000") = vbNo Then
                bSetFocus Me, "txtSales"
                Exit Function
            End If
        End If
        
        TillFieldsOK = True
    
    Else
        MsgBox "Please Select an Item from the List"
    End If

End Function

Public Function InitTillItem()

    labelTillDescription.Caption = ""
    txtSales.Text = ""
    txtSalesDP.Text = ""
    labelSales1Price.Caption = ""
    labelSalesDPPrice.Caption = ""
    
    txtGlass.Text = ""
    txtGlassDP.Text = ""
    labelGlass1Price.Caption = ""
    labelGlassDPPrice.Caption = ""
    
    lblGlass.Visible = False
    lblGlassDP.Visible = False
    labelGlass1Price.Visible = False
    labelGlassDPPrice.Visible = False
    txtGlass.Visible = False
    txtGlassDP.Visible = False
    
    
    
    'DP
    lblSalesDP.Visible = bDualPrice
    txtSalesDP.Visible = bDualPrice
    labelSalesDPPrice.Visible = bDualPrice
    
    
    cmdTillSave.Enabled = False
    lTillID = 0


End Function

Public Function ClearClientLastCountFigures(lID As Long)
Dim rs As Recordset
Dim rsDel As Recordset

    On Error GoTo ErrorHandler
    
    ' Here we Delete the last delivieries record and clear out the last stock take figures
    ' for the passed client ID
    
    Set rs = SWdb.OpenRecordset("SELECT ID, FullQty, Open, Weight, RcvdQty, SalesQty, SalesQtyDP, GlassQty, GlassQtyDP FROM tblClientProductPLUs WHERE ClientID = " & Trim$(lID))
    ' Here we get all the product/plu IDs associated with this client
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        Do
            
            SWdb.Execute "Delete * FROM tblDeliveries WHERE ClientProdPLUID = " & Trim$(rs("ID") + 0)
            ' This deletes the corresponding Deliveries record
            
            rs.Edit
            rs("FullQty") = Empty
            rs("Open") = Empty
            rs("weight") = Empty
            rs("RcvdQty") = Empty
            rs("SalesQty") = Empty
            rs("SalesQtyDP") = Empty
            rs("GlassQty") = Empty
            rs("GlassQtyDP") = Empty
            
            rs.Update
            ' This clears out the figures from the last stock take
        
            rs.MoveNext
        Loop While Not rs.EOF
    
    End If
    
    ClearClientLastCountFigures = True

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("ClearClientLastCountFigures") Then Resume 0
    Resume CleanExit

End Function

Public Function SaveLastStockFigures(lID As Long)
Dim rs As Recordset
Dim rsUpd As Recordset

    On Error GoTo ErrorHandler
    
    Set rsUpd = SWdb.OpenRecordset("tblClientProductPLUs")
    rsUpd.Index = "ID"
    
    
    Set rs = SWdb.OpenRecordset("SELECT SUM(FullQty) AS lFullQty, SUM(Open) AS lOpen FROM tblClientProductPLUs INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID WHERE ClientID = " & Trim$(lID))
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Set rs = SWdb.OpenRecordset("SELECT * FROM tblClientProductPLUs INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID WHERE ClientID = " & Trim$(lID))
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst
    
            Do
                
                rsUpd.Seek "=", rs("tblClientProductPLUs.ID")
                If Not rsUpd.NoMatch Then
                    rsUpd.Edit
                    rsUpd("LastQty") = CalcAmount(rs("FullQty"), rs("Open"), rs("Weight"), rs("txtFullWeight"), rs("txtEmptyWeight"))
                    rsUpd.Update
                    rsUpd.Bookmark = rsUpd.LastModified
                End If
            
                rs.MoveNext
            Loop While Not rs.EOF
        End If
    
    End If
    
    rsUpd.Close
    rs.Close
    ' here check to see if figures to be saved are all zero
        
    SaveLastStockFigures = True

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rsUpd Is Nothing Then Set rsUpd = Nothing
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("SaveLastStockFigures") Then Resume 0
    Resume CleanExit



End Function

Public Function GetClientDate(lCLId As Long, lDtID) As Boolean
Dim rs As Recordset
            
    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("SELECT * FROM tblDates WHERE ClientID = " & Trim$(lCLId) & " ORDER By InProgress, From DESC", dbOpenSnapshot)
    If Not (rs.BOF And rs.EOF) Then
        
        labelFrom.Caption = Format(rs("From"), "dd mmm yy (ddd)")
        labelTo.Caption = Format(rs("To"), "dd mmm yy (ddd)")
        labelDate.Caption = Format(rs("On"), "dd mmm yy")
        labelFrom.Tag = Format(rs("From"), sDMY)
        labelTo.Tag = Format(rs("To"), sDMY)
                
        lDtID = rs("ID") + 0
        
    Else
        labelFrom.Caption = ""
        labelTo.Caption = ""
        labelFrom.Tag = ""
        labelTo.Tag = ""
    
    End If
    
    GetClientDate = True
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetClientDate") Then Resume 0
    Resume CleanExit

End Function

Public Function SetBeginEndStockTakeButtons(lSelClID As Long)

'VER1.4

    If CountInProgress(lSelClientID) Then
    
    ' For any count in progress allow date edit
    ' count going on and same client selected
    
        grdMenu.Cell(flexcpText, 1, 1) = "Edit Stock Take Dates"
        cmdStart.Caption = "Save New Dates"
        
        cmdCancel.Visible = Not bEvaluation ' only show cancel if its not an evaluation
        
        grdMenu.RowHidden(21) = False
        grdMenu.Height = 22 * grdMenu.RowHeight(0)
        ' show the 'X' button
    
    Else
    ' otherwise allow another stock take to begin...
        
        grdMenu.Cell(flexcpText, 1, 1) = "Begin New Stock Take"
        cmdStart.Caption = "Begin Stock Take"
        cmdCancel.Visible = False
        grdMenu.RowHidden(21) = True
        grdMenu.Height = 21 * grdMenu.RowHeight(0)
        
        ' hide the 'X' button
    
    End If

    SetBeginEndStockTakeButtons = True

End Function


Public Sub ShowStockTakeSymbol()

    If CountInProgress(lSelClientID) Then
    
        imgInProgress.Visible = True
    Else
        imgInProgress.Visible = False
    End If

End Sub

Public Function ConfirmClientName()
Dim sTmp As String
Dim lTempClientID As Long
Dim lTempDatesID As Long
    
    
    If Not CountInProgress(lSelClientID) Then
    ' there is no count in progress for this client
        
        If GetCountsInProgress(lTempClientID, lTempDatesID) > 0 Then
        ' and there are counts in progress... so warn
        
            If MsgBox("Warning - Stock Take already in progress for other Client, Continue with changes to client: " & Replace(lblClient.Tag, "_", " "), vbDefaultButton2 + vbYesNo + vbQuestion, "Modifying records for Client " & Replace(lblClient.Tag, "_", " ")) = vbYes Then
                ConfirmClientName = True
            End If
        
        Else
            ConfirmClientName = True
        
        End If
        
    Else
        ConfirmClientName = True

    End If
    
End Function

Public Function SetSymbolMsgMenu(lSelId As Long) As Boolean
Dim iRow As Integer

    ' AUDIT IN PROGRESS SYMBOL / MESSAGE / COLOUR
    
' VER1.4

    If CountInProgress(lSelClientID) Then
    
        labelStatus = "Count In Progress"
        labelStatus.ForeColor = sDarkRed
        imgInProgress.Visible = True
    
    Else
    
        labelStatus = "Last Stock Take"
        labelStatus.ForeColor = sDarkGreen
        imgInProgress.Visible = False
    
    End If
    
    If IsDate(labelFrom.Tag) And IsDate(labelTo.Tag) Then
    
         For iRow = 1 To grdMenu.Rows - 1
            grdMenu.Cell(flexcpForeColor, iRow, 0, iRow, 1) = sBlack
         Next
         grdMenu.Tag = True
            
    ElseIf lDatesID > 0 Then
   
         grdMenu.Cell(flexcpForeColor, 1, 0, 1, 1) = sBlack
         
         For iRow = 2 To grdMenu.Rows - 1
            grdMenu.Cell(flexcpForeColor, iRow, 0, iRow, 1) = sLightGrey
         Next
        
         grdMenu.Tag = False
    
    Else    ' new client name selected but it doesnt have any audit dates
   
         For iRow = 2 To grdMenu.Rows - 1
            grdMenu.Cell(flexcpForeColor, iRow, 0, iRow, 1) = sLightGrey
         Next
        
         grdMenu.Tag = False
    
    End If


    SetSymbolMsgMenu = True

End Function

Public Sub SetCtrlButtons(bhow As Boolean)

    cmdStockTake.Enabled = bhow
    cmdProducts.Enabled = bhow
    cmdReports.Enabled = bhow
    cmdEmail.Enabled = bhow
End Sub
Public Sub ShowMenu(bhow As Boolean)

    'picStatus.Height = Me.Height - 360 - picSelect.Height
    picStatus.Visible = bhow

End Sub

Public Sub SetupFrameAndGrid(sObj As String, sTag As String, sTitle As String, bPrintFrameVisible As Boolean, bGridVisible As Boolean)

'ver531
            grdCount.FontSize = 10

    labelTitle.Caption = sTitle
    labelTitle.Tag = sTag
    
    Select Case sTag
    
        Case 1, 2, 3, 4, 5, 6
    
        frmStockWatch.Controls(sObj).Visible = bPrintFrameVisible
             
        grdCount.Visible = bGridVisible
        grdCount.ScrollBars = flexScrollBarVertical

        frmStockWatch.Controls(sObj).Left = picStatus.Width + (Me.Width - picStatus.Width - frmStockWatch.Controls(sObj).Width) / 2
        grdCount.Left = frmStockWatch.Controls(sObj).Left + 120
    
        frmStockWatch.Controls(sObj).Top = picSelect.Height + picSelect.Top + 100
        grdCount.Top = frmStockWatch.Controls(sObj).Top + 3020
    
        grdCount.Width = fraStock.Width - 240
        frmStockWatch.Controls(sObj).Height = Me.Height - (picSelect.Top + picSelect.Height) - 1760
        grdCount.Height = frmStockWatch.Controls(sObj).Height - (picSelect.Top + picSelect.Height) - 1760
    
    
        Case Else
    
                
                frmStockWatch.Controls(sObj).Visible = bPrintFrameVisible
                     
                grdCount.Visible = bGridVisible
                grdCount.ScrollBars = flexScrollBarVertical
        
                frmStockWatch.Controls(sObj).Left = picStatus.Width
        '        grdCount.Left = frmStockWatch.Controls(sObj).Left + 120
        '
                frmStockWatch.Controls(sObj).Top = picSelect.Height + picSelect.Top
        '        grdCount.Top = frmStockWatch.Controls(sObj).Top + 3020
        '
                frmStockWatch.Controls(sObj).Width = Me.Width - picStatus.Width
        '        grdCount.Width = fraStock.Width - 240
        '
                frmStockWatch.Controls(sObj).Height = Me.Height - (picSelect.Top + picSelect.Height)
        '        grdCount.Height = frmStockWatch.Controls(sObj).Height - (picSelect.Top + picSelect.Height) - 1760
            
    
    End Select
    
    DoEvents

End Sub

Public Function getProjectedSales()

    labelProjectedSales = Val(labelCalculatedSales.Caption) - _
                            (Val(txtStaff) + _
                            Val(txtComplimentary) + _
                            Val(txtWastage) + _
                            Val(txtOverRings) + _
                            Val(txtPromotions) + _
                            Val(txtOffLicense) + _
                            Val(txtVoucherSales) + _
                            Val(txtKitchen) + _
                            Val(txtOther))

End Function

Public Sub SetupShowPrevious(bTill As Boolean)

 '   cboShowPrevious.ListIndex = 0
        
    lblShowPrevious.Visible = bTill
    cboShowPrevious.Visible = bTill
    

End Sub

Public Function DoSave(iDelivery As Integer, iRef As Integer)
Dim iSaveTopRow As Integer

  If ConfirmClientName() Then
    
    If DeliveryFieldsOK() Then
    
        cmdDelSave.Enabled = False
'        cmdAddNew.Enabled = False
        
        bHourGlass True
        
        If SaveDelivery(lDelID, lNextDelID, iDelivery, iRef) Then
          
'Ver 2.2.0
          If UpdatePurchasePrice(lNextDelID) Then
'=========
            
'Ver 2.4

            If iDelivery = 1 And ((iRef = 1) Or (iRef = 2 And Not bEditOther)) Then
            ' A new record added or a record is deleted so show all deliveries again
                    
                    iSaveTopRow = grdCount.TopRow
                    
                    gbOk = ShowDeliveries()
                    ' reshow deliveries since we may have added a new one
            
                    grdCount.TopRow = iSaveTopRow
'==========
            
            Else

                gbOk = UpdateDelivery(lDelID, lNextDelID)
            End If
            
            LogMsg Me, "Delivery: " & labelDelItem.Caption & " Saved for " & Replace(lblClient.Tag, "_", " "), labelDelItem.Caption & " Qty:" & txtQty & " free:" & txtFree & " Cost:" & txtDelCost

            labelDelCount.Caption = SetCount(grdCount, "Quantity")
            
            gbOk = InitDeliveryItem()
                
            If GetNextDelivery(lNextDelID) Then

                bSetFocus Me, "txtQty"
            Else
                bSetFocus Me, "grdMenu"
            End If
            
          End If
          
        End If
    
        cmdDelSave.Enabled = True
'        cmdAddNew.Enabled = True
        
        bHourGlass False
        '
    
    End If
  End If

  bHourGlass False

End Function

Public Function AddUp(sSum As String)
Dim itotal As Integer

    ' sSum = 3+2400 + 300+ 120
    
    Do
        If InStr(sSum, "+") > 0 Then
            itotal = itotal + Val(Left(sSum, InStr(sSum, "+") - 1))
            sSum = Trim$(Mid(sSum, InStr(sSum, "+") + 1, Len(sSum)))
        Else
            itotal = itotal + Val(sSum)
            sSum = ""
        End If
        
    Loop While sSum <> "" And sSum <> "+"
    
    AddUp = Trim$(itotal)
    
End Function

Public Function SetBoxWidth(sChars As String)

        If (Len(sChars) + 1) * 153 > iBoxWidth Then
        ' use openitem box for reference
        
            SetBoxWidth = (Len(sChars) * 153)
        
        Else
            SetBoxWidth = iBoxWidth
        End If

End Function

Public Function SetBoxColoursWhite()

    txtActual.BackColor = vbWhite
    txtStaff.BackColor = vbWhite
    txtComplimentary.BackColor = vbWhite
    txtWastage.BackColor = vbWhite
    txtOverRings.BackColor = vbWhite
    txtPromotions.BackColor = vbWhite
    txtOffLicense.BackColor = vbWhite
    txtVoucherSales.BackColor = vbWhite
    txtKitchen.BackColor = vbWhite
    txtOther.BackColor = vbWhite
    txtSurplus.BackColor = vbWhite
    
End Function

Public Sub HideAllFrames()

    Me.Picture = imgList.ListImages("bkg").Picture
    
    fraPrint.Visible = False
'    frmCtrl.Visible = False
    picCash.Visible = False
'    picSummary.Visible = False
    fraSales.Visible = False
    fraStock.Visible = False
    fraDates.Visible = False
    fraDelivery.Visible = False
    
    Cal.Visible = False
    grdCount.Visible = False
    'grdList.Visible = False
    grdClients.Visible = False
End Sub

Public Function GetStockID(iCode As Integer)
Dim rs As Recordset
    
    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("SELECT tblProducts.*, tblClientProductPLUs.ID, tblClientProductPLUs.ClientID FROM (tblProducts INNER JOIN tblClientProductPLUs ON tblProducts.ID = tblClientProductPLUs.ProductID) WHERE ClientID = " & Trim$(lSelClientID) & " AND txtCode = '" & Trim$(iCode) & "'", dbOpenSnapshot)

    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        GetStockID = rs("tblClientProductPLUs.ID") + 0
    
    End If
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetStockID") Then Resume 0
    Resume CleanExit

End Function

Public Sub SetupView(sWhich As String)

    frmCtrl.grdList.Visible = True
    frmCtrl.grdList.Rows = 1
    ' controls needed - made visible
    
    SetUpActiveList sMenuCtrl
    ' Set up Active/Inactive Cbo List
         
'    frmCtrl.Show vbModal
    

End Sub

Public Function SetupControl(sWhich As String)


    With frmCtrl

    'ver 3.0.6 (9) Move form up to reveal bottom of screen
    
    .Left = (Screen.Width - .Width) / 2
    .Top = 100
    '-----------------------------------------------------
    
    .cboActive.Left = 1920
    .lblActive.Left = 780

    .optProduct.Visible = False
    .optPLU.Visible = False
    .cboActive.Clear
    .lblByGroup.Visible = False
    .cboByGroup.Visible = False
    .cmdNew.Enabled = True

    Select Case sWhich

        Case "Clients"
         .fraSearch.Visible = False
'         picSelect.Visible = False

         .cboActive.AddItem "Active Clients"
         .cboActive.AddItem "InActive Clients"
         .cboActive.Visible = True
         .lblActive.Caption = "&View"

        Case "Products"
         .fraSearch.Visible = True
'         picSelect.Visible = False
'picStatus.Visible = False

         .cboActive.AddItem "Active Products"
         .cboActive.AddItem "InActive Products"
         .cboActive.Visible = True
         .lblActive.Caption = "&View"

        Case "PLUs"
         .fraSearch.Visible = True
'         picSelect.Visible = False

         .cboActive.AddItem "Active Products"
         .cboActive.AddItem "InActive Products"
         .cboActive.Visible = True
         .lblActive.Caption = "&View"

        Case "Groups"
         .fraSearch.Visible = False
'         picSelect.Visible = False

         .cboActive.AddItem "Active Groups"
         .cboActive.AddItem "InActive Groups"
         .cboActive.Visible = True
         .lblActive.Caption = "&View"

         .cboActive.Left = 8730
         .lblActive.Left = 7560
         .optProduct.Visible = True
         .optPLU.Visible = True

        Case "Product/PLUs"
         .lblByGroup.Visible = True
         .cboByGroup.Visible = True

         .fraSearch.Visible = False
'         picSelect.Visible = False

         .cboActive.AddItem "Client PLUs and Stock Products"
         .cboActive.AddItem "Client PLUs"
         .cboActive.AddItem "Client Stock Products"
         .cboActive.Visible = True
         .lblActive.Caption = "&View"

         .grdList.BackColorAlternate = sLightGreen

         picSelect.Visible = True
'         fraCtrl.Top = picSelect.Height
'         frmCtrl.Visible = True


    End Select
    End With



End Function

Public Function RestoreLastCountFigures(lID As Long)
Dim rs As Recordset
Dim rsDel As Recordset

    On Error GoTo ErrorHandler
    
    ' Here we Delete the last delivieries record and copy back the Last Quantities to the Full Quantities
    
    Set rs = SWdb.OpenRecordset("SELECT ID, FullQty, LastQty, Open, Weight, RcvdQty, SalesQty, SalesQtyDP, GlassQty, GlassQtyDP FROM tblClientProductPLUs WHERE ClientID = " & Trim$(lID))
    ' Here we get all the product/plu IDs associated with this client
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        Do
            
            SWdb.Execute "Delete * FROM tblDeliveries WHERE ClientProdPLUID = " & Trim$(rs("ID") + 0)
            ' This deletes the corresponding Deliveries record
            
            rs.Edit
            rs("FullQty") = rs("LastQty")
            rs("Open") = 0
            rs("weight") = 0
            rs("RcvdQty") = 0
            rs("SalesQty") = 0
            rs("SalesQtyDP") = 0
            rs("GlassQty") = 0
            rs("GlassQtyDP") = 0
            
            rs.Update
            ' This clears out the figures from the last stock take
        
            rs.MoveNext
        Loop While Not rs.EOF
    
    End If
    
    RestoreLastCountFigures = True

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("RestoreLastCountFigures") Then Resume 0
    Resume CleanExit


End Function

Public Function GetCountsInProgress(lID As Long, lDtID As Long) As Integer
Dim rs As Recordset

    On Error GoTo ErrorHandler

    ' This looks for counts in Progress
    ' If there's only one then return IDs or if > 1 then just return the number in progress

    Set rs = SWdb.OpenRecordset("SELECT * FROM tblDates WHERE InProgress = true", dbOpenSnapshot)
    ' first make sure its not there already

    If Not (rs.EOF And rs.BOF) Then

        rs.MoveLast
        
        If rs.RecordCount = 1 Then
            rs.MoveFirst
            
            lID = rs("ClientID") + 0
            lDtID = rs("ID") + 0
            GetCountsInProgress = 1
        
        Else
            rs.MoveLast
''            lID = 0
''            lDtID = 0
            GetCountsInProgress = rs.RecordCount + 0
        End If

    Else
        lID = 0
        lDtID = 0
        GetCountsInProgress = 0

    End If

    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks

    If Not rs Is Nothing Then Set rs = Nothing
'

    Exit Function

ErrorHandler:
    If CheckDBError("CountsInProgress") Then Resume 0
    Resume CleanExit


End Function

Public Function SetCountInProgress(lID As Long, bhow As Boolean)

     If grdClients.FindRow(lID) > -1 Then
        grdClients.Cell(flexcpData, grdClients.FindRow(lID), 0) = bhow
        SetCountInProgress = True
    End If

End Function

Public Function ClearInProgress(lID As Long)

'ver530
    SWdb.Execute "UPDATE tblDates SET InProgress = false, CountStep = 0 WHERE ID = " & Trim$(lID)

End Function

Public Sub BlockAll(bhow As Boolean)

    grdMenu.Enabled = Not bhow
    picSelect.Enabled = Not bhow
    
    btnClients.Enabled = Not bhow
    btnProducts.Enabled = Not bhow
    btnPlus.Enabled = Not bhow
    btnGroups.Enabled = Not bhow
    btnSettings.Enabled = Not bhow
    btnAbout.Enabled = Not bhow
    btnAudits.Enabled = Not bhow
    
    imgReport.Visible = Not bhow
    
End Sub

Public Function UpdatePurchasePrice(lID As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler

    Set rs = SWdb.OpenRecordset("tblClientProductPLUs")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lID
    If Not rs.NoMatch Then
        If Val(txtDelCost) <> Val(rs("PurchasePrice")) Then
            rs.Edit
            rs("PurchasePrice") = Val(txtDelCost)
            rs.Update
        End If
    
    End If
    
    UpdatePurchasePrice = True
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks

    If Not rs Is Nothing Then Set rs = Nothing
'

    Exit Function

ErrorHandler:
    If CheckDBError("UpdatePurchasePrice") Then Resume 0
    Resume CleanExit

End Function

Public Function DeleteDeliveryItem(iRow As Integer)
Dim rs As Recordset

    On Error GoTo ErrorHandler

    Set rs = SWdb.OpenRecordset("tblDeliveries")
    rs.Index = "PrimaryKey"
    If grdCount.Cell(flexcpData, iRow, 0) > 0 Then
        rs.Seek "=", grdCount.Cell(flexcpData, iRow, 0)
        If Not rs.NoMatch Then
            rs.Delete
        End If

        DeleteDeliveryItem = True
        
    End If
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks

    If Not rs Is Nothing Then Set rs = Nothing

    Exit Function

ErrorHandler:
    If CheckDBError("DeleteDeliveryItem") Then Resume 0
    Resume CleanExit

End Function


Public Sub SetDeliveryLocked(bhow As Boolean)

    cboDelivery.Locked = bhow
    
    If bhow Then
        cboDelivery.BackColor = &H8000000F
    Else
        cboDelivery.BackColor = vbWhite
    End If
End Sub

Public Sub SetFreeLocked(bhow As Boolean)
    
    txtFree.Locked = bhow
    txtDelCost.Locked = bhow
    
    If bhow Then
        txtFree.Text = ""
        txtFree.BackColor = &H8000000F
        txtDelCost.BackColor = &H8000000F
    Else
        txtFree.BackColor = vbWhite
        txtDelCost.BackColor = vbWhite
    End If

End Sub

Public Sub SetEditButton(bhow As Boolean)

    
    bEditOther = bhow

    If bhow Then
        cmdDelSave.Caption = "&Save"
    Else
        cmdDelSave.Caption = "&Delete"
    End If
End Sub


Public Sub InitStockWatch()

    lSelClientID = 0
    lblClient.Caption = ""
    lblClient.Tag = ""

    SetCtrlButtons vbFalse
    ' Enable buttons
    
    imgInProgress.Visible = False
    picStatus.Visible = False
    Me.Picture = imgList.ListImages("bkg").Picture

End Sub

Public Function SaveNoteAndProductsTicked(lID As Long)
Dim rs As Recordset
Dim sTickedproducts As String
'ver 310
' SummaryNote filed now contains the summary note and the products ticked
' e.g. /01/04/11/13/17//summary note here//
    
    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblDates")
    rs.Index = "PrimaryKey"
    
    rs.Seek "=", lID
    If Not rs.NoMatch Then
    
        'ver 310
        sTickedproducts = SaveProductsTicked()
        
        rs.Edit
        
        'ver 310
        rs("SummaryNote") = sTickedproducts & "~" & Trim$(txtNote.Text)
        
        rs.Update
    
    End If
    
    rs.Close
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("SaveNoteAndProductsTicked") Then Resume 0
    Resume CleanExit

End Function

Public Function OnDateCheck(sOn As String)

    If IsDate(sOn) Then
        
        If DateValue(sOn) <= DateAdd("d", 1, Format(Now, sDMY)) Then
        
            OnDateCheck = True
    
        End If
    End If

End Function

Public Function GetInvoiceName(lDtsID As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblDates")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lDtsID
    If Not rs.NoMatch Then
        GetInvoiceName = gbRegion & "_" & rs("INVNumber")
    End If
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks

    If Not rs Is Nothing Then Set rs = Nothing
    '

    Exit Function

ErrorHandler:
    If CheckDBError("GetInvoiceName") Then Resume 0
    Resume CleanExit

End Function

Public Function InvoiceCreated(lDtsID As Long)
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    InvoiceCreated = False  ' default
    
    Set rs = SWdb.OpenRecordset("tblDates")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lDtsID
    If Not rs.NoMatch Then
        If Not IsNull(rs("INVNumber")) Then
            If Val(rs("InvNumber")) > 0 Then
                InvoiceCreated = True
            End If
        End If
    End If
    
    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks

    If Not rs Is Nothing Then Set rs = Nothing
    '

    Exit Function

ErrorHandler:
    If CheckDBError("InvoiceCreated") Then Resume 0
    Resume CleanExit


End Function

Public Sub GoCreateInvoice()

'    timSummaryFiles.Enabled = False
        
    frmInvoice.lDatesID = lDatesID
    frmInvoice.lClientID = lSelClientID

    frmInvoice.Show vbModal
    ' View invoice

'    timSummaryFiles.Enabled = True

End Sub

Public Function SetFormSize()

    Me.Width = Screen.Width
    Me.Height = Screen.Height - 520
    Me.Top = 0
    Me.Left = 0

End Function

Public Function SaveProductsTicked()
Dim iRow As Integer
Dim sTicked As String

    ' loop all products on list
    ' select those that are ticked
    ' add them to the string
    
    ' return string
    
    For iRow = grdCount.FixedRows To grdCount.Rows - 1
    
        If grdCount.Cell(flexcpChecked, iRow, grdCount.ColIndex("Sel")) = flexChecked Then
            sTicked = sTicked & "|" & grdCount.RowData(iRow)
        End If
        
    Next
    
    If sTicked <> "" Then SaveProductsTicked = sTicked & "|"
    
End Function

Public Function totalSelected()
Dim iRow As Integer
Dim curTotal As Long
    
    For iRow = 2 To grdCount.Rows - 1
    
        If grdCount.Cell(flexcpChecked, iRow, grdCount.ColIndex("Sel")) = flexChecked Then
        
            curTotal = curTotal + Val(grdCount.Cell(flexcpTextDisplay, iRow, grdCount.ColIndex("Value")))
        
        End If
        
    Next
    
' Ver 553
    ' Added getPercentOfActual function in Ver 553
    
    labelTotal = "Total: " & Format(curTotal, "Currency") & " (" & GetPercentOfActual(curTotal) & ")"

End Function

Public Function GetPercentOfActual(curTot As Long) As String
Dim rs As Recordset

' Ver 553 - NEW FUNCTION

    On Error GoTo ErrorHandler
    
    curTot = Abs(curTot)
    
    If curTot <> 0 Then
    
        Set rs = SWdb.OpenRecordset("tblDates")
        rs.Index = "PrimaryKey"
        rs.Seek "=", lDatesID
        
        If rs("Actual") > 0 Then
        
            If Not rs.NoMatch Then
                GetPercentOfActual = Format((curTot / rs("Actual")) * 100, "0.00") & "%"
            End If
        End If
        rs.Close
    Else
        GetPercentOfActual = "0"
    
    End If
    
Leave:
    Exit Function
    
ErrorHandler:

    MsgBox Trim$(Error)
    Resume Leave

End Function

Public Function SelectedMax()
Dim iRow As Integer
Dim iSel As Integer

    For iRow = 2 To grdCount.Rows - 1
    
        If grdCount.Cell(flexcpChecked, iRow, grdCount.ColIndex("Sel")) = flexChecked Then
    
            iSel = iSel + 1
        End If
        
    Next

    If iSel < 25 Then SelectedMax = True

End Function

' ver 440 Removed for now...to preserve memory
'Public Function CheckForNewAgentProgram()
'Dim sNewRev As String
'Dim sFile, sFilesCol, sf
'Dim sFileName As String
'Dim objfile
'
'    Set objfile = CreateObject("Scripting.FileSystemObject")
'    ' create object
'
'    ' NOW CHECK TO SEE IF THERE'S A NEW SWIAgent PROGRAM.
'
'    If objfile.FileExists(sDBLoc & "/" & "SWIAgentNEW.exe") Then
'    ' Check for StockWatchNEW.exe
'
'        On Error Resume Next
'
'        Dim Process As Object
'        For Each Process In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = 'SWIAgent.exe'")
'            Process.Terminate
'        Next
'        ' stop SWIAgent if its running
'
'        sFileName = sDBLoc & "/" & "SWIAgentOLD.exe"
'        Set sFile = objfile.GetFile(sFileName)
'        sf = sFile.Delete
'        ' kill old
'
'        ' RENAME PREV AS OLD
'
'        Name sDBLoc & "/" & "SWIAgentPREV.exe" As sDBLoc & "/" & "SWIAgentOLD.exe"
'        ' rename prev to old
'
'        ' RENAME CUR AS PREV
'
'        Name sDBLoc & "/" & "SWIAgent.exe" As sDBLoc & "/" & "SWIAgentPREV.exe"
'        ' rename current to prev
'
'        ' RENAME NEW VER TO STOCKWATCH.EXE
'
'        Name sDBLoc & "/" & "SWIAgentNEW.exe" As sDBLoc & "/" & "SWIAgent.exe"
'        ' rename New to current
'
'        gbOk = RestartAgentProgram()
'        ' restart
'
'        CheckForNewAgentProgram = True
'    End If
'
'Leave:
'
'    Exit Function
'
'NoOldFile:
'    Resume Next
'
'NoNewFile:
'    Print lAUDf, "No New File: " & sDBLoc & "/" & "StockwatchNEW.exe"
'
'NoPrevFileToRename:
'    Resume Next
'
'NoCurFileToRename:
'    Print #lAUDf, "No Current file: " & sDBLoc & "/" & "Stockwatch.exe"
'    Resume Next
'
'NoCopy:
'    Print #lAUDf, "New file not copied: " & sDBLoc & "/" & "StockwatchVer " & sNewRev & ".exe"
'    Resume Leave
'
'
'End Function

Public Function GetStockWatchVersion(sFile As String)
Dim objfile As Object
Dim sNewRev As String

    Set objfile = CreateObject("Scripting.filesystemObject")
    sNewRev = objfile.GetFileVersion(sFile)
    sNewRev = Replace(sNewRev, ".", "")
    GetStockWatchVersion = Left(sNewRev, 2) & Right(sNewRev, 1)

End Function

Public Sub ShowSelected()
Dim iRow As Integer

    For iRow = 2 To grdCount.Rows - 2
        If grdCount.Cell(flexcpChecked, iRow, grdCount.ColIndex("Sel")) = 1 Then
            grdCount.RowHidden(iRow) = False
        Else
            grdCount.RowHidden(iRow) = True
        End If
    
    Next
    
End Sub

Public Function UpdateTillSalesQuantities(lID As Long, sSales1 As String, sSales2 As String)
Dim rs As Recordset

    Set rs = SWdb.OpenRecordset("tblClientProductPLUs")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lID
    If Not rs.NoMatch Then
    
        rs.Edit
        rs("SalesQty") = Val(sSales1)
        rs("SalesQtyDP") = Val(sSales2)
        rs.Update
    
    End If
    
    rs.Close

    UpdateTillSalesQuantities = True

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks

    If Not rs Is Nothing Then Set rs = Nothing
    '

    Exit Function

ErrorHandler:
    If CheckDBError("UpdateTillSalesQuantities") Then Resume 0
    Resume CleanExit

End Function

Public Function UpdateQtyAllItemsSamePLU(lID As Long, sPLU As String, sSales As String, sSalesDP As String, sGlass As String, sGlassDP As String)
Dim rs As Recordset
Dim rsUpd As Recordset

    ' All new function for ver433
    
    
    Set rs = SWdb.OpenRecordset("SELECT * FROM tblClientProductPLUs WHERE ClientId = " & Trim$(lID) & " AND PLUNumber = " & sPLU)
    If Not (rs.EOF And rs.BOF) Then
    
        rs.MoveFirst
    
        Set rsUpd = SWdb.OpenRecordset("tblClientProductPLUs")
        rsUpd.Index = "PrimaryKey"
        
        Do
            rsUpd.Seek "=", rs("ID")
            If Not rsUpd.NoMatch Then
            
            'ver530
            ' The original SalesQty field is kept
            ' The value of Pints is now added to the number of glasses (divided by the measure)
            ' for draught its 2 for wines its 4
                
                rsUpd.Edit
                
'                If iGlass > 0 Then
'                    rsUpd("SalesQty") = Val(sSales) + (Val(sGlass) / iGlass)
 '                   rsUpd("SalesQtyDP") = Val(sSalesDP) + (Val(sGlassDP) / iGlass)
'                    rsUpd("SalesQty") = Val(sSales)
'                    rsUpd("SalesQtyDP") = Val(sSalesDP)
                
'                Else
                    rsUpd("SalesQty") = Val(sSales)
                    rsUpd("SalesQtyDP") = Val(sSalesDP)
'                End If
                
                rsUpd("GlassQty") = Val(sGlass)
                rsUpd("GlassQtyDP") = Val(sGlassDP)
                rsUpd.Update
            
            End If
    
            rs.MoveNext
        Loop While Not rs.EOF
    
    End If
    
    rs.Close
    rsUpd.Close
    
    UpdateQtyAllItemsSamePLU = True

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks

    If Not rs Is Nothing Then Set rs = Nothing
    '

    Exit Function

ErrorHandler:
    If CheckDBError("UpdateQtyAllItemsSamePLU") Then Resume 0
    Resume CleanExit


End Function
Public Function GetBarList()
Dim rs As Recordset

    ' Ver 440 this is a new function to return a list of
    ' bars for a particular client
    
    cboBar.Clear
    cboBar.AddItem "All Bars"
    cboBar.ItemData(cboBar.NewIndex) = 0
    
    
    Set rs = SWdb.OpenRecordset("SELECT * FROM tblbars WHERE ClientID = " & Trim$(lSelClientID))
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        Do
            cboBar.AddItem rs("Bar") & ""
            cboBar.ItemData(cboBar.NewIndex) = rs("ID") + 0
            rs.MoveNext
        Loop While Not rs.EOF
        
        GetBarList = True
        
    End If
    rs.Close
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    If Not rs Is Nothing Then Set rs = Nothing
    bHourGlass False
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetBarList ") Then Resume 0
    Resume CleanExit

    
End Function

Public Function ClearClientLastBarFigures(lID As Long)

    On Error GoTo CleanExit
    
    ' Here we Delete the last bar figures recorded for the passed client ID
    
    If lID > 0 Then
    
        SWdb.Execute "Delete * FROM tblBarCount WHERE ClientID = " & Trim$(lID)
            
    End If
    
    ClearClientLastBarFigures = True

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
'    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("ClearClientLastBarFigures") Then Resume 0
    Resume CleanExit

End Function

Public Sub ShowBarComboList(bhow As Boolean)

    lblBar.Visible = bhow
    cboBar.Visible = bhow

    
'    btnAdd.Visible = bhow
    
End Sub

Public Sub InitEnableInput(bhow As Boolean)

    txtFullQty.Enabled = bhow
    txtOpen.Enabled = bhow
    txtWeight.Enabled = bhow
    cmdSaveStock.Enabled = bhow
    
End Sub

Public Function SaveBarCount(lStkId As Long, lClientID As Long, lBrID As Long)
Dim rs As Recordset

    ' This is a new function in Ver440
    ' This saves the count of an individual item for a particular bar
    ' If its already counted (and entered) then its just edited.
   
    If lBrID > 0 Then
    
        Set rs = SWdb.OpenRecordset("SELECT * FROM tblBarCount WHERE ClientProdPluID = " & Trim$(lStkId) & " AND BarID = " & Trim$(lBrID))
        If Not (rs.EOF And rs.BOF) Then
        ' First see can we edit it ...
        
            rs.Edit
            
        Else
        ' If not it hasnt been saved yet ... so save it now
            rs.AddNew
            rs("ClientProdPluID") = lStkId
            rs("ClientID") = lSelClientID
            rs("BarID") = lBrID
            rs("BarFullQty") = Empty
            rs("BarOpen") = Empty
            rs("BarWeight") = Empty
            
        End If
        
        rs("BarFullQty") = Val(txtFullQty.Text)
        
        If lblFullQty.Tag = "True" Then
                rs("BarOpen") = Val(txtOpen.Text)
                rs("BarWeight") = Val(txtWeight.Text)
        End If

        rs.Update
        rs.Close
    
    End If
    
    SaveBarCount = True
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    If Not rs Is Nothing Then Set rs = Nothing
    bHourGlass False
    
    Exit Function

ErrorHandler:
    If CheckDBError("SaveBarCount ") Then Resume 0
    Resume CleanExit

End Function

Public Function ShowBarStockCount(lBrID As Long)
Dim rs As Recordset
Dim rsBar As Recordset
Dim iLastGroup As Integer
Dim iRow As Integer

    On Error GoTo ErrorHandler
    
    ' Ver440
    ' This is all new function to get bar stock figures.
    ' The first SQL is just to display products, groups descriptions very like the
    ' ShowTotalStockCount Routine.
    
    ' The 2nd SQL is to display the counts from the table tblBarCount
    
    grdCount.Rows = 1
    grdCount.Cols = 0
    
    ' Use Count ID to get Till Sales so far
    
    SetupCountField frmStockWatch, "Code", ""
    SetupCountField frmStockWatch, "Description", ""
    SetupCountField frmStockWatch, "Size", ""
    SetupCountField frmStockWatch, "Full Qty", ""
    SetupCountField frmStockWatch, "Open Items", ""
    SetupCountField frmStockWatch, "Weight", ""

'ver440 out for now
'    SetupCountField frmStockWatch, "Chk", ""
    
    grdCount.ScrollBars = flexScrollBarVertical
    
    ' FIRST SQL JUST LOADS UP GRID WITH PRODUCTS BUT WITHOUT THE QUANTITES
    
    
    Set rs = SWdb.OpenRecordset("SELECT txtCode, tblClientProductPLUs.ID, cboGroups, tblProductGroup.txtDescription, tblProducts.txtDescription, txtSize FROM (tblClientProductPLUs INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE tblClientProductPLUs.ClientID = " & Trim$(lSelClientID) & " AND tblProducts.chkActive = true AND tblClientPRoductPLUs.Active = true ORDER BY tblProducts.cboGroups, tblProducts.txtDescription, txtSize", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
    
        rs.MoveFirst
        Do
            
            If iLastGroup <> rs("cboGroups") Then
    
                grdCount.AddItem vbTab & rs("cboGroups") & "  " & rs("tblProductGroup.txtDescription")
                iLastGroup = rs("cboGroups")
                grdCount.Cell(flexcpAlignment, grdCount.Rows - 1, 1) = flexAlignLeftCenter
                grdCount.Cell(flexcpBackColor, grdCount.Rows - 1, 0, grdCount.Rows - 1, grdCount.Cols - 1) = &HC0C0C0
                grdCount.Cell(flexcpForeColor, grdCount.Rows - 1, 0, grdCount.Rows - 1, grdCount.Cols - 1) = sBlack
    
            End If
            
            grdCount.AddItem rs("txtCode") & vbTab & rs("tblProducts.txtDescription") & vbTab & rs("txtSize")
            grdCount.Cell(flexcpForeColor, grdCount.Rows - 1, 0, grdCount.Rows - 1, grdCount.Cols - 1) = sDarkGrey
            grdCount.RowData(grdCount.Rows - 1) = rs("ID") + 0
                
            rs.MoveNext
        Loop While Not rs.EOF
    
        End If
        
    SetColWidths Me, "grdCount", "Description", False
    
    labelStockCount.Caption = SetCount(grdCount, "FullQty")
    
    
    ' THIS SQL GETS THE TOTALS FOR THE BAR SELECTED AND ADD THEM ONTO THE GRID
        
    Set rs = SWdb.OpenRecordset("SELECT * FROM tblBarCount WHERE ClientID = " & Trim$(lSelClientID) & " AND BarID = " & Trim$(lBrID), dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
    
        rs.MoveFirst
        
        Do
        
            iRow = grdCount.FindRow(rs("ClientProdPLUID") + 0)
            If iRow > 0 Then
            
                If IsNull(rs("BarFullQty")) Then
                    grdCount.Cell(flexcpForeColor, iRow, 0, iRow, grdCount.Cols - 1) = sDarkGrey

                ElseIf rs("BarFullQty") = Empty And rs("BarOpen") = Empty Then
                    grdCount.Cell(flexcpText, iRow, grdCount.ColIndex("fullqty"), iRow, grdCount.ColIndex("weight")) = "0"
                    grdCount.Cell(flexcpForeColor, iRow, 0, iRow, grdCount.Cols - 1) = sBlack

                Else
                    grdCount.Cell(flexcpText, iRow, grdCount.ColIndex("fullqty")) = rs("BarFullQty")
                    grdCount.Cell(flexcpText, iRow, grdCount.ColIndex("openitems")) = rs("BarOpen")
                    grdCount.Cell(flexcpText, iRow, grdCount.ColIndex("weight")) = rs("BarWeight")
                    grdCount.Cell(flexcpForeColor, iRow, 0, iRow, grdCount.Cols - 1) = sBlack

                End If
                
'ver440 out for now
'                If rs("Verified") Then
'                    grdCount.Cell(flexcpData, iRow, grdCount.ColIndex("Chk")) = True
'                    grdCount.Cell(flexcpPicture, iRow, grdCount.ColIndex("Chk")) = imgList.ListImages("tick").Picture
'                Else
'                    grdCount.Cell(flexcpData, iRow, grdCount.ColIndex("Chk")) = False
'
'                End If
            
            End If
        
            rs.MoveNext
        
        Loop While Not rs.EOF

    End If
    
'    grdCount.Cell(flexcpPictureAlignment, 1, grdCount.ColIndex("Chk"), grdCount.Rows - 1, grdCount.ColIndex("Chk")) = 4
    
    labelStockCount.Caption = SetCount(grdCount, "FullQty")
    
    ShowBarStockCount = True

    rs.Close

CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("ShowBarStockCount") Then Resume 0
    Resume CleanExit

    


End Function

Public Function ShowBarCountsForThisItem(lID As Long)
Dim rs As Recordset

    ' Ver 440 new function for POP-UP Panel
    
    ' This shows the counts in the different bars for the same item.
    ' They are shown on a pop up grid when the line item is clicked.
    
    
    Set rs = SWdb.OpenRecordset("SELECT BarID, BarFullQty, BarOpen, BarWeight FROM tblBarCount WHERE ClientProdPLUID = " & Trim$(lID) & " AND ClientID = " & Trim$(lSelClientID) & " ORDER BY BarID ASC", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
    
        rs.MoveFirst
        
        grdBars.Rows = 0
        
        Do
            grdBars.AddItem GetBar(rs("BarID")) & vbTab & rs("BarFullQty") & vbTab & rs("BarOpen") & vbTab & rs("BarWeight")
            rs.MoveNext
        Loop While Not rs.EOF
    
        grdBars.Height = grdBars.Rows * grdBars.RowHeight(0)
        picBars.Height = grdBars.Height + 420

        ShowBarCountsForThisItem = True
    
    End If


CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("ShowBarCountsForThisItem") Then Resume 0
    Resume CleanExit


End Function

Public Function GetBar(lID As Long)
Dim rs As Recordset

    ' Ver 440 returns the bar description
    
    Set rs = SWdb.OpenRecordset("tblBars")
    rs.Index = "PrimaryKey"
    rs.Seek "=", lID
    If Not rs.NoMatch Then
    
        GetBar = rs("Bar") & ""
    End If

    rs.Close
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    If Not rs Is Nothing Then Set rs = Nothing
    '
    
    Exit Function

ErrorHandler:
    If CheckDBError("GetBar") Then Resume 0
    Resume CleanExit



End Function

'ver440 out for now
'Public Function SaveTick(lStkId As Long, lBrID As Long, bhow As Boolean)
'Dim rs As Recordset
'
'    ' This is a new function in Ver440
'    ' This saves the Count Tick when doing a count verify
'    ' bHow will toggle it on or off
'
'    Set rs = SWdb.OpenRecordset("SELECT [Verified] FROM tblBarCount WHERE ClientProdPluID = " & Trim$(lStkId) & " AND BarID = " & Trim$(lBrID))
'    If Not (rs.EOF And rs.BOF) Then
'    ' First see can we edit it ...
'
'        rs.Edit
'        rs("Verified") = bhow
'        rs.Update
'    End If
'
'    rs.Close
'
'    SaveTick = True
'
'CleanExit:
'    'DBEngine.Idle dbRefreshCache
'    ' Release unneeded DB locks
'    If Not rs Is Nothing Then Set rs = Nothing
'    bHourGlass False
'
'    Exit Function
'
'ErrorHandler:
'    If CheckDBError("SaveTick ") Then Resume 0
'    Resume CleanExit
'
'End Function

Public Function CheckAndApplyAnyUpdate()
Dim TblDef As TableDef
Dim fld As Field

' new in ver440 . This is for applying database changes

' This just checks to see if one of the fields is present that are supposed to be there
' with the update and if not apply the full update

' This check and its function (DoDatabaseUpdate) can be removed on the next minor update

' If there is a future update requiring a database change it can be added in here again
    
    
    
'------------
''    Set TblDef = SWdb.TableDefs("tblClients")
''    For Each fld In TblDef.Fields
''        If fld.Name = "chkMultipleBars" Then
''            CheckAndApplyAnyUpdate = True
''            Exit Function
''        End If
''    ' Check for field if its not there apply the full update
''
''    Next
''
''    gbOk = DoDatabaseUpdate()
''    ' Update wasnt applied so do it here now
'''------------
    CheckAndApplyAnyUpdate = True
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    bHourGlass False
    
    Exit Function

ErrorHandler:
    If CheckDBError("CheckAndApplyAnyUpdate ") Then Resume 0
    Resume CleanExit

End Function

Public Function GenerateTabletOutputFile(sDrv As String, lID As Long)
Dim rs As Recordset
Dim lOutf As Long
Dim sFileName As String

    ' New function with ver 500
    
    lOutf = FreeFile
    sFileName = sDrv & "SWCountFile.csv"
    Open sFileName For Output As #lOutf
    ' open output file
    
    ' CLIENT NAME AND ADDRESS
    
    Set rs = SWdb.OpenRecordset("SELECT ID, txtName, rtfAddress FROM tblClients WHERE ID = " & Trim$(lID))
    If Not (rs.EOF And rs.BOF) Then
    ' Here we get the client ID, client name and client address
    
        Print #lOutf, "!Client" & "," & Trim$(rs("ID") + 0) & "," & rs("txtName") & "," & Replace(rs("rtfAddress"), vbCrLf, " ")
    End If
    rs.Close
    
    ' AUDIT FROM TO DATES
    
    Set rs = SWdb.OpenRecordset("SELECT [On],[from],[to] FROM tblDates WHERE ClientID = " & Trim$(lID) & " AND InProgress = true")
    If Not (rs.EOF And rs.BOF) Then
    ' Here we get the client ID, client name and client address
    
        Print #lOutf, "!Dates" & "," & rs("On") & "," & rs("From") & "," & rs("To")
    End If
    rs.Close
    
    
    ' CLIENTS BARS
    
    Set rs = SWdb.OpenRecordset("SELECT * FROM tblBars WHERE ClientID = " & Trim$(lID))
    ' Each bar ID and bar description for the client
    If Not (rs.EOF And rs.BOF) Then
    
        rs.MoveFirst
        Do
            Print #lOutf, "!Bar" & "," & Trim$(rs("ID") + 0) & "," & rs("Bar")
            rs.MoveNext
        Loop While Not rs.EOF
    End If
    rs.Close
    
    
    ' ALL ITEMS
    
    'ver 522
    Set rs = SWdb.OpenRecordset("SELECT * FROM (tblClientProductPLUs INNER JOIN tblProducts ON tblClientProductPLUs.ProductID = tblProducts.ID) INNER JOIN tblProductGroup ON tblProducts.cboGroups = tblProductGroup.ID WHERE tblClientProductPLUs.ClientID = " & Trim$(lID) & " AND Active = true ORDER BY tblProducts.cboGroups, tblProducts.txtDescription, txtSize", dbOpenSnapshot)
    ' SQl return list of plus for this clinet sorted by plu#
    
    If Not (rs.EOF And rs.BOF) Then

        rs.MoveFirst
        Do
            Print #lOutf, "!Item" & "," & rs("tblClientProductPLUs.ID") & "," & rs("cboGroups") & "  " & rs("tblProductGroup.txtDescription") & "," & rs("txtCode") & "," & rs("tblProducts.txtDescription") & "," & rs("txtSize") & "," & rs("txtFullWeight") & "," & rs("txtEmptyWeight") & "," & rs("chkOpenItem") & "," & rs("txtIssueUnits")
            rs.MoveNext
        Loop While Not rs.EOF
    
    End If
    Print #lOutf, "!END"
    Close #lOutf
    
    rs.Close
    
    
    MsgBox sFileName & " Created"
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    bHourGlass False
    
    Exit Function

ErrorHandler:
    If CheckDBError("GenerateTabletOutputFile ") Then Resume 0
    Resume CleanExit

End Function


Public Sub SetMeasureBoxes(iGls As Integer)

    If iGls > 0 Then
        lblGlass.Visible = True
        txtGlass.Visible = True
        labelGlass1Price.Visible = True
        
        If bDualPrice Then
            lblGlassDP.Visible = True
            txtGlassDP.Visible = True
            labelGlassDPPrice.Visible = True
        End If
        
    Else
        lblGlass.Visible = False
        txtGlass.Visible = False
        labelGlass1Price.Visible = False
        
        lblGlassDP.Visible = False
        txtGlassDP.Visible = False
        labelGlassDPPrice.Visible = False
        
    
    End If


End Sub

Public Function ClearValuation(lID As Long)

    SWdb.Execute "DELETE from tblDates WHERE ID = " & Trim$(lID)
    
CleanExit:
    'DBEngine.Idle dbRefreshCache
    ' Release unneeded DB locks
    
    Exit Function

ErrorHandler:
    If CheckDBError("ClearValuation ") Then Resume 0
    Resume CleanExit


End Function

Public Function SaveSummaryAndNote(bhow As Boolean)

        btnSaveNote.Enabled = bhow

End Function

Public Sub RemoveHiddenRows()
Dim iRow As Integer

    For iRow = grdCount.Rows - 1 To 2 Step -1
        If grdCount.RowHidden(iRow) = True Then
        
            grdCount.RemoveItem (iRow)
        
        End If
    
    Next

End Sub

