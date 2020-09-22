VERSION 5.00
Object = "*\AsubformCTL\Seven.vbp"
Begin VB.Form frmForm2 
   Caption         =   "Form2"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin Seven.SubForm SubForm1 
      Height          =   4020
      Left            =   120
      TabIndex        =   0
      Top             =   255
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   7091
      BackColor       =   -2147483633
      BorderStyle     =   1
      Moveable        =   -1  'True
      Sizeable        =   -1  'True
      Caption         =   "frmForm2 with frmTip subform"
   End
End
Attribute VB_Name = "frmForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Me.SubForm1.AttachForm frmTip
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Me.SubForm1.DetachForm
End Sub
