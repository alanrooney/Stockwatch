VERSION 5.00
Begin VB.Form frmInfoMsg 
   BorderStyle     =   0  'None
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmInfoMsg.frx":0000
   ScaleHeight     =   795
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblMsg 
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
      ForeColor       =   &H0025213F&
      Height          =   255
      Left            =   750
      TabIndex        =   0
      Top             =   270
      Width           =   3645
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmInfoMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dbDelay As Double

Private Sub Form_Activate()
Dim X As Integer

        endofpause# = Timer + dbDelay
    Do
        X% = DoEvents()
    Loop While Timer < endofpause#

    Unload Me


    ' 30 limit of message length



End Sub

