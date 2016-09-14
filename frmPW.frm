VERSION 5.00
Begin VB.Form frmPW 
   BorderStyle     =   0  'None
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   ControlBox      =   0   'False
   Icon            =   "frmPW.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPW.frx":1CCA
   ScaleHeight     =   915
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPass 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1350
      PasswordChar    =   "~"
      TabIndex        =   0
      Top             =   270
      Width           =   1395
   End
   Begin VB.Label lblCount 
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
      Height          =   255
      Left            =   390
      TabIndex        =   1
      Top             =   330
      Width           =   2115
   End
End
Attribute VB_Name = "frmPW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then Unload Me

End Sub

Private Sub Form_Load()

    bPassGood = False
    ' reset to make sure

End Sub

Private Sub Txtpass_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        bPassGood = False
        
        If txtPass = "sw2011" Then
        
            bPassGood = True
        End If
    
        Unload Me
    
    End If

End Sub

