VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   0  'None
   Caption         =   "StockWatch Settings"
   ClientHeight    =   11805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13665
   Icon            =   "frmSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11805
   ScaleWidth      =   13665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MyCommandButton.MyButton btnClose 
      Height          =   255
      Left            =   6360
      TabIndex        =   0
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
   Begin VB.PictureBox picEmail 
      BorderStyle     =   0  'None
      Height          =   5835
      Left            =   3420
      Picture         =   "frmSettings.frx":1CCA
      ScaleHeight     =   5835
      ScaleWidth      =   6750
      TabIndex        =   41
      Top             =   3000
      Visible         =   0   'False
      Width           =   6750
      Begin VB.TextBox txtSWEmail 
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
         Left            =   3090
         TabIndex        =   55
         Top             =   3975
         Width           =   2655
      End
      Begin VB.TextBox txtPassword 
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
         Left            =   3075
         TabIndex        =   51
         Top             =   2955
         Width           =   2655
      End
      Begin VB.TextBox txtUsername 
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
         Left            =   3090
         TabIndex        =   49
         Top             =   2460
         Width           =   2655
      End
      Begin VB.CheckBox chkSSL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F1EDEF&
         Caption         =   "SSL"
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
         Left            =   5025
         TabIndex        =   47
         Top             =   1980
         Width           =   705
      End
      Begin VB.TextBox txtPort 
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
         Left            =   3105
         TabIndex        =   46
         Top             =   1950
         Width           =   645
      End
      Begin VB.TextBox txtfrom 
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
         Left            =   3075
         TabIndex        =   53
         Top             =   3480
         Width           =   2655
      End
      Begin VB.TextBox txtsmtp 
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
         Left            =   3105
         TabIndex        =   44
         Top             =   1485
         Width           =   2655
      End
      Begin MyCommandButton.MyButton btnOk 
         Height          =   495
         Left            =   5610
         TabIndex        =   56
         Top             =   5010
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
         Caption         =   "Ok"
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
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Stockwatch Email Address"
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
         Left            =   645
         TabIndex        =   54
         Top             =   4020
         Width           =   2595
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Watch - Email Settings"
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
         Left            =   225
         TabIndex        =   42
         Top             =   150
         Width           =   2550
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Left            =   2100
         TabIndex        =   50
         Top             =   3000
         Width           =   1995
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "UserName"
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
         Left            =   2040
         TabIndex        =   48
         Top             =   2490
         Width           =   1995
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Port"
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
         Left            =   2670
         TabIndex        =   45
         Top             =   2040
         Width           =   1995
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "From Email Address"
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
         TabIndex        =   52
         Top             =   3510
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Outgoing Mail (SMTP)"
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
         Left            =   1095
         TabIndex        =   43
         Top             =   1515
         Width           =   1995
      End
   End
   Begin MyCommandButton.MyButton btnQuit 
      Height          =   495
      Left            =   5415
      TabIndex        =   1
      Top             =   4950
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
   Begin MyCommandButton.MyButton btn 
      Height          =   360
      Index           =   0
      Left            =   285
      TabIndex        =   2
      Top             =   735
      Width           =   1545
      _ExtentX        =   2725
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
      AppearanceMode  =   1
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   13748165
      BackColorDisabled=   13748165
      BorderColor     =   32768
      BorderDrawEvent =   1
      BorderWidth     =   0
      TransparentColor=   14215660
      Caption         =   "&Print Defaults"
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
   Begin MyCommandButton.MyButton btn 
      Height          =   360
      Index           =   1
      Left            =   1830
      TabIndex        =   3
      Top             =   735
      Width           =   1545
      _ExtentX        =   2725
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
      AppearanceMode  =   1
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   13748165
      BackColorDisabled=   13748165
      BorderColor     =   32768
      BorderDrawEvent =   1
      BorderWidth     =   0
      TransparentColor=   14215660
      Caption         =   "&Vat Rates"
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
   Begin MyCommandButton.MyButton btn 
      Height          =   360
      Index           =   2
      Left            =   3375
      TabIndex        =   4
      Top             =   735
      Width           =   1545
      _ExtentX        =   2725
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
      AppearanceMode  =   1
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   13748165
      BackColorDisabled=   13748165
      BorderColor     =   32768
      BorderDrawEvent =   1
      BorderWidth     =   0
      TransparentColor=   14215660
      Caption         =   "&Bank Details"
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
   Begin MyCommandButton.MyButton btn 
      Height          =   360
      Index           =   3
      Left            =   4920
      TabIndex        =   5
      Top             =   735
      Width           =   1545
      _ExtentX        =   2725
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
      AppearanceMode  =   1
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   13748165
      BackColorDisabled=   13748165
      BorderColor     =   32768
      BorderDrawEvent =   1
      BorderWidth     =   0
      TransparentColor=   14215660
      Caption         =   "&Invoice Text"
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
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   5820
      Index           =   1
      Left            =   6900
      Picture         =   "frmSettings.frx":66EC
      ScaleHeight     =   5820
      ScaleWidth      =   6750
      TabIndex        =   10
      Top             =   5925
      Width           =   6750
      Begin VB.CheckBox chkVatActive 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   4995
         TabIndex        =   15
         Top             =   3465
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00DDDDDD&
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
         Left            =   945
         MaxLength       =   1
         TabIndex        =   12
         Top             =   3435
         Width           =   645
      End
      Begin VB.TextBox txtRate 
         BackColor       =   &H00DDDDDD&
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
         Left            =   1605
         MaxLength       =   5
         TabIndex        =   13
         Top             =   3435
         Width           =   1125
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H00DDDDDD&
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
         Left            =   2775
         MaxLength       =   30
         TabIndex        =   14
         Top             =   3435
         Width           =   1995
      End
      Begin VSFlex8LCtl.VSFlexGrid grdVat 
         Height          =   1725
         Left            =   885
         TabIndex        =   11
         Top             =   1695
         Width           =   5025
         _cx             =   8864
         _cy             =   3043
         Appearance      =   1
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
         BackColorSel    =   9929356
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   16053492
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483639
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   8
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
         FormatString    =   $"frmSettings.frx":B587
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
      Begin MyCommandButton.MyButton cmdVatAdd 
         Default         =   -1  'True
         Height          =   495
         Left            =   3480
         TabIndex        =   16
         Top             =   3990
         Width           =   2400
         _ExtentX        =   4233
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
         Caption         =   "&Add / Update Vat Rates"
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Watch - Vat Rates"
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
         Left            =   210
         TabIndex        =   37
         Top             =   120
         Width           =   2160
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code     Rate (%)               Description                 Active"
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
         Left            =   1065
         TabIndex        =   17
         Top             =   1380
         Width           =   4515
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   5820
      Index           =   2
      Left            =   15
      Picture         =   "frmSettings.frx":B5F8
      ScaleHeight     =   5820
      ScaleWidth      =   6750
      TabIndex        =   22
      Top             =   0
      Width           =   6750
      Begin VB.TextBox txtBIC 
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
         Height          =   300
         Left            =   2205
         MaxLength       =   11
         TabIndex        =   33
         ToolTipText     =   "Bank Account Number"
         Top             =   3120
         Width           =   3555
      End
      Begin VB.TextBox txtIBAN 
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
         Height          =   300
         Left            =   2205
         MaxLength       =   22
         TabIndex        =   35
         ToolTipText     =   "Bank Sort code"
         Top             =   3510
         Width           =   3555
      End
      Begin VB.TextBox txtName 
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
         Height          =   300
         Left            =   2190
         MaxLength       =   100
         TabIndex        =   27
         ToolTipText     =   "Name of Bank Account Holder"
         Top             =   1920
         Width           =   3555
      End
      Begin VB.TextBox txtSortCode 
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
         Height          =   300
         Left            =   2205
         MaxLength       =   10
         TabIndex        =   31
         ToolTipText     =   "Bank Sort code"
         Top             =   2715
         Width           =   3555
      End
      Begin VB.TextBox txtAccountNo 
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
         Height          =   300
         Left            =   2205
         MaxLength       =   12
         TabIndex        =   29
         ToolTipText     =   "Bank Account Number"
         Top             =   2325
         Width           =   3555
      End
      Begin VB.TextBox txtBank 
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
         Height          =   300
         Left            =   2205
         MaxLength       =   50
         TabIndex        =   25
         ToolTipText     =   "Bank Name"
         Top             =   1485
         Width           =   3555
      End
      Begin MyCommandButton.MyButton btnSaveBankDetails 
         Height          =   495
         Left            =   5400
         TabIndex        =   36
         Top             =   3945
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
         Caption         =   "&Save"
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
      Begin VB.Label lblBIC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BIC"
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
         Left            =   1800
         TabIndex        =   32
         Top             =   3150
         Width           =   315
      End
      Begin VB.Label lblIBAN 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IBAN"
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
         Left            =   1650
         TabIndex        =   34
         Top             =   3540
         Width           =   465
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Details will only appear on Invoice"
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
         Height          =   315
         Left            =   1170
         TabIndex        =   39
         Top             =   4035
         Width           =   3825
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
         Height          =   375
         Left            =   2175
         TabIndex        =   38
         Top             =   1455
         Visible         =   0   'False
         Width           =   3630
      End
      Begin VB.Label lblSortCode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sort Code"
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
         Left            =   1215
         TabIndex        =   30
         Top             =   2745
         Width           =   900
      End
      Begin VB.Label lblAccountNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account No"
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
         Left            =   1080
         TabIndex        =   28
         Top             =   2355
         Width           =   1035
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name On Account"
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
         Left            =   480
         TabIndex        =   26
         Top             =   1950
         Width           =   1620
      End
      Begin VB.Label lblBank 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
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
         Left            =   1650
         TabIndex        =   24
         Top             =   1515
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Watch - Bank Details"
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
         Left            =   210
         TabIndex        =   23
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   5820
      Index           =   3
      Left            =   6885
      Picture         =   "frmSettings.frx":1001A
      ScaleHeight     =   5820
      ScaleWidth      =   6750
      TabIndex        =   18
      Top             =   0
      Width           =   6750
      Begin VB.TextBox txtDefaultInvoice 
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
         Height          =   1680
         Left            =   540
         TabIndex        =   19
         Top             =   1635
         Width           =   5640
      End
      Begin MyCommandButton.MyButton btnSaveText 
         Height          =   495
         Left            =   5400
         TabIndex        =   20
         Top             =   3945
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
         Caption         =   "&Save"
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
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "This is the default text to appear on an Invoice."
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
         Left            =   915
         TabIndex        =   40
         Top             =   4050
         Width           =   4245
      End
      Begin VB.Label lblDefaultInvoice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Watch - Invoice Text"
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
         Left            =   210
         TabIndex        =   21
         Top             =   120
         Width           =   2355
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   5820
      Index           =   0
      Left            =   0
      Picture         =   "frmSettings.frx":1486A
      ScaleHeight     =   5820
      ScaleWidth      =   6750
      TabIndex        =   6
      Top             =   5925
      Width           =   6750
      Begin VSFlex8LCtl.VSFlexGrid grdReps 
         Height          =   2565
         Left            =   1380
         TabIndex        =   7
         Top             =   1665
         Width           =   3975
         _cx             =   7011
         _cy             =   4524
         Appearance      =   1
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
         BackColorSel    =   9929356
         ForeColorSel    =   16711680
         BackColorBkg    =   -2147483644
         BackColorAlternate=   16053492
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483639
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   8
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   9
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSettings.frx":19655
         ScrollTrack     =   0   'False
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   0.1
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
         BackColorFrozen =   16777215
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblPrintDefaults 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Watch - Print Defaults"
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
         Left            =   210
         TabIndex        =   9
         Top             =   120
         Width           =   2445
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report                           Select"
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
         Left            =   2565
         TabIndex        =   8
         Top             =   1395
         Width           =   2400
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lVatRateID As Long
Public bFormIsShown As Boolean

Private Sub btn_Click(Index As Integer)
    
    pic(0).Visible = False
    pic(1).Visible = False
    pic(2).Visible = False
    pic(3).Visible = False
    
    btn(0).ToggleValue = False
    btn(1).ToggleValue = False
    btn(2).ToggleValue = False
    btn(3).ToggleValue = False
    
    btn(Index).ToggleValue = True
    
    pic(Index).Visible = True

'    If Index = 0 Then bSetFocus Me, "grdReps"


End Sub

Private Sub btnClose_Click()

    btnQuit_Click

End Sub

Private Sub btnOk_Click()
                
    SaveSetting appname:=App.Title, Section:="Email", Key:="smtp", Setting:=txtsmtp
    SaveSetting appname:=App.Title, Section:="Email", Key:="port", Setting:=txtPort
    SaveSetting appname:=App.Title, Section:="Email", Key:="from", Setting:=txtfrom

    SaveSetting appname:=App.Title, Section:="Email", Key:="ssl", Setting:=chkSSL
    SaveSetting appname:=App.Title, Section:="Email", Key:="username", Setting:=txtUsername
    SaveSetting appname:=App.Title, Section:="Email", Key:="password", Setting:=txtPassword
    SaveSetting appname:=App.Title, Section:="Email", Key:="swEmail", Setting:=txtSWEmail
    
    GetEmailDefaults

    picEmail.Visible = False

End Sub

Private Sub btnQuit_Click()

    Unload Me

End Sub

Private Sub btnSaveBankDetails_Click()

'    If Bankfieldsok() Then
    
        
        If SaveBankInfo() Then

            LogMsg frmStockWatch, "Bank Details Added/Updated ", ""
            
            MsgBox "Bank Details Added/Updated Ok"
        
        End If
        
'    End If
    

End Sub

Private Sub btnSaveText_Click()

    gbOk = SaveDefaultInvoiceText()

End Sub

Private Sub chkVatActive_Click()
    cmdVatAdd.Enabled = True

End Sub

Private Sub cmdVatAdd_Click()
    
    If lVatRateID = 0 And grdVat.FindRow(txtCode, , 0) > 0 Then
        MsgBox "This Tax Code already In Use, Please use another one"
        bSetFocus Me, "txtCode"
        Exit Sub
    
    ElseIf txtCode <> "" And txtRate <> "" Then
        
        gbOk = WriteDB(Me, "Vat", lVatRateID, False, 4, _
                            txtCode, _
                            txtRate, _
                            txtDescription, _
                            chkVatActive)
        
        LogMsg frmStockWatch, "vat Rates Modified " & txtCode, "Rate:" & txtRate & " Desc:" & txtDescription
        
         gbOk = ReadDB(Me, "Vat", 0, 4, _
                            txtCode, _
                            txtRate, _
                            txtDescription, _
                            chkVatActive)
   
        If gbOk Then MsgBox "Vat Rate Updated Ok"
        InitVatEntry
        
    Else
        MsgBox "Please enter Vat Code/Vat Rate"
    End If
    
    bSetFocus Me, "txtCode"


End Sub

Private Sub Form_Activate()
    
    
    If Not bFormIsShown Then
    
        gbOk = WindowAppear(Me, 0, (Screen.Height - Me.Height) / 2, (Screen.Width - Me.Width) / 2, 3, False)
        
        bFormIsShown = True
    End If

    bSetFocus Me, "grdReps"
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    bFormIsShown = False

End Sub

Private Sub grdReps_Click()

'    If grdReps.Col = 1 Then grdReps.Col = 0

'    setClearSelection grdReps.Row, grdReps.Col

End Sub

Private Sub grdReps_KeyPress(KeyAscii As Integer)
    
'    If grdReps.Row > -1 Then
        
        If grdReps.FindRow(UCase(Chr$(KeyAscii)), , 0) > -1 Then
        
            grdReps.Row = grdReps.FindRow(UCase(Chr$(KeyAscii)), , 0)
        
            setClearSelection grdReps.Row, 2
        
        
        End If

'    End If

End Sub

Private Sub grdReps_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    setClearSelection grdReps.Row, grdReps.Col
    
End Sub

Private Sub grdVat_Click()


  If grdVat.Row > -1 Then
    
      lVatRateID = grdVat.RowData(grdVat.Row)
      gbOk = ReadDB(Me, "Vat", lVatRateID, 4, _
                            txtCode, _
                            txtRate, _
                            txtDescription, _
                            chkVatActive)
    bSetFocus Me, "txtCode"
  End If
  
'  Else
'    InitVatEntry
'  End If

End Sub

Private Sub Form_Load()
Dim sBank As String
Dim sNameOnAccount As String
Dim sAccountNo As String
Dim sSortCode As String
Dim sBIC As String
Dim sIBAN As String


        gbOk = SetupPics()
         
        gbOk = ReadDB(Me, "Vat", 0, 4, _
                            txtCode, _
                            txtRate, _
                            txtDescription, _
                            chkVatActive)
            
         
        txtDefaultInvoice.Text = GetDefaultInvoiceText()
            
        If GetBankInfo(sBank, sNameOnAccount, sAccountNo, sSortCode, sBIC, sIBAN) Then
            txtBank = sBank
            txtName = sNameOnAccount
            txtAccountNo = sAccountNo
            txtSortCode = sSortCode
            txtBIC = sBIC
            txtIBAN = sIBAN
        End If
        
        gbOk = SetupMenu()
         
        cmdVatAdd.Enabled = False

        btn_Click (0)

End Sub

Private Sub MyButton1_Click()

End Sub

Private Sub txtAccountNo_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtAccountNo_KeyPress(KeyAscii As Integer)
    KeyAscii = CharOk(KeyAscii, 0, " ") ' 0 = no only, 1 = char only, 2 = both

End Sub

Private Sub txtAccountNo_LostFocus()
    lblAccountNo.ForeColor = sBlack

End Sub

Private Sub txtBank_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtBank_KeyPress(KeyAscii As Integer)
    KeyAscii = CharOk(KeyAscii, 2, "-_/ &.'") ' 0 = no only, 1 = char only, 2 = both

End Sub

Private Sub txtBank_LostFocus()
    lblBank.ForeColor = sBlack

End Sub

Private Sub txtBIC_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtBIC_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub txtBIC_LostFocus()
    lblBIC.ForeColor = sBlack

End Sub

Private Sub txtCode_Change()
    cmdVatAdd.Enabled = True
End Sub

Private Sub txtDescription_Change()
    cmdVatAdd.Enabled = True

End Sub

Private Sub txtIBAN_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtIBAN_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub txtIBAN_LostFocus()
    lblIBAN.ForeColor = sBlack

End Sub

Private Sub txtName_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    KeyAscii = CharOk(KeyAscii, 2, "-_/ &.'") ' 0 = no only, 1 = char only, 2 = both

End Sub

Private Sub txtName_LostFocus()
    lblName.ForeColor = sBlack

End Sub

Private Sub txtRate_Change()
    cmdVatAdd.Enabled = True

End Sub
Private Sub InitVatEntry()
  txtCode.Text = ""
  txtRate.Text = ""
  txtDescription = ""
  chkVatActive.Value = 1
  lVatRateID = 0
  
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    ' return pressed force focus to next available object in tabbing order
        gbOk = GotoNextControl(Me, 0)
    
    ElseIf KeyAscii = 27 Then
        Unload Me
        
    End If


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
    grdReps.RowData(7) = "H"
    grdReps.RowData(8) = "I"
    
    
    For iRow = 0 To 8
        grdReps.Cell(flexcpPicture, iRow, 0) = frmStockWatch.imgList.ListImages("button").Picture
    
        If GetSetting(App.Title, "Reports", App.Title & "Print " & grdReps.RowData(iRow)) = "-1" Then
            grdReps.Cell(flexcpChecked, iRow, 2) = True
        Else
            grdReps.Cell(flexcpChecked, iRow, 2) = False
        End If
    
    Next

End Function

Public Sub setClearSelection(iRow As Integer, iCol As Integer)
Dim sRep As String
    
    If iCol = 2 Then
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

Public Function SaveDefaultInvoiceText()
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    Set rs = SWdb.OpenRecordset("tblInvoiceDefaults")
    rs.Index = "PrimaryKey"
    If rs.EOF Then
        rs.AddNew
    Else
        rs.MoveFirst
        rs.Edit
    End If
    
    rs("InvoiceText") = txtDefaultInvoice
    rs.Update
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
        If CheckDBError("SaveDefaultInvoiceText") Then Resume 0
        Resume CleanExit
    End If


End Function


Public Function SetupPics()
Dim iCnt As Integer

    For iCnt = 0 To 3
    
        pic(iCnt).Left = 0
        pic(iCnt).Top = 0
        
        pic(iCnt).Visible = False
    
    Next
    
    picEmail.Left = 0
    picEmail.Top = 0
    picEmail.Visible = False
    
    Me.Width = 6750
    Me.Height = 5820

End Function

Private Sub txtSortCode_GotFocus()
    gbOk = bSetupControl(Me)

End Sub

Private Sub txtSortCode_KeyPress(KeyAscii As Integer)
    KeyAscii = CharOk(KeyAscii, 0, "- ") ' 0 = no only, 1 = char only, 2 = both

End Sub

Private Sub txtSortCode_LostFocus()
    lblSortCode.ForeColor = sBlack

End Sub

Public Function Bankfieldsok()

    If Trim$(txtBank) <> "" Then
        If Trim$(txtName) <> "" Then
            If Trim$(txtAccountNo) <> "" Then
                If Trim$(txtSortCode) <> "" Then
                    If Trim$(txtBIC) <> "" Then
                        If Trim$(txtIBAN) <> "" Then
                        
                            Bankfieldsok = True
                    
                        Else
                            MsgBox "Please Enter IBAN"
                            bSetFocus Me, "txtIBAN"
                        End If
                    Else
                        MsgBox "Please Enter BIC"
                        bSetFocus Me, "txtBIC"
                    End If
                Else
                    MsgBox "Please Enter Bank Sort Code"
                    bSetFocus Me, "txtSortCode"
                End If
            Else
                MsgBox "Please Enter Bank Account Number"
                bSetFocus Me, "txtAccountNo"
            End If
        Else
            MsgBox "Please Enter Name of Account Holder"
            bSetFocus Me, "txtName"
        End If
    Else
        MsgBox "Please Enter Bank Name"
        bSetFocus Me, "txtBank"
    End If
    
End Function
Public Function SaveBankInfo()
Dim sBnk As String
Dim rs As Recordset

    On Error GoTo ErrorHandler
    
    sBnk = "<Bank>" & Trim$(txtBank.Text) & "/<Bank>" & _
        "<NameOnAccount>" & Trim$(txtName.Text) & "/<NameOnAccount>" & _
        "<AccountNo>" & Trim$(txtAccountNo.Text) & "/<AccountNo>" & _
        "<SortCode>" & Trim$(txtSortCode.Text) & "/<SortCode>" & _
        "<BIC>" & Trim$(txtBIC.Text) & "/<BIC>" & _
        "<IBAN>" & Trim$(txtIBAN.Text) & "/<IBAN>"
    
    Set rs = SWdb.OpenRecordset("tblFranchisee")
    rs.Index = "PrimaryKey"
    If Not rs.EOF Then
        rs.MoveFirst
        rs.Edit
    
        ' ENCRYPTED BANK DETAILS
        rs("BankInfo") = Encrypt(sBnk, sKey)
        ' rebundle
    
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
    If CheckDBError("SaveBankInfo") Then Resume 0
    Resume CleanExit

End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     
    If KeyCode = Asc("M") Then
        If Shift = 2 Then
            
            If picEmail.Visible Then
                picEmail.Visible = False
            Else
                picEmail.Visible = True
            
                txtsmtp = gbSMTP
                txtPort = gbPort
                txtfrom = gbEmailfromAddress
            
                chkSSL = gbSSL
                txtUsername = gbUsername
                txtPassword = gbPassword
                txtSWEmail = gbSWEmail

            End If
        End If
    End If

End Sub
