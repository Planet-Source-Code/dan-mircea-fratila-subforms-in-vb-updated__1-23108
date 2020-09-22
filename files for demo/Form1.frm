VERSION 5.00
Object = "*\AsubformCTL\Seven.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin Seven.SubForm SubForm2 
      Height          =   2715
      Left            =   5355
      TabIndex        =   18
      Top             =   5025
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   4789
      BorderStyle     =   1
   End
   Begin Seven.SubForm SubForm1 
      Height          =   3690
      Left            =   180
      TabIndex        =   17
      Top             =   75
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   6509
      ScrollBarStyle  =   1
      BackColor       =   -2147483633
      BorderStyle     =   1
      Picture         =   "Form1.frx":0000
      Moveable        =   -1  'True
      Sizeable        =   -1  'True
      Caption         =   "Custom Caption HERE"
      Begin VB.TextBox Text1 
         Height          =   645
         Left            =   1935
         TabIndex        =   19
         Text            =   "Other Controls"
         Top             =   1050
         Width           =   1365
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   870
         Left            =   2310
         Shape           =   5  'Rounded Square
         Top             =   75
         Width           =   990
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "<- Caption"
      Height          =   540
      Left            =   7695
      TabIndex        =   16
      Top             =   3960
      Width           =   1035
   End
   Begin VB.TextBox Text2 
      Height          =   525
      Left            =   5340
      TabIndex        =   15
      Text            =   "Custom Caption HERE"
      Top             =   3960
      Width           =   2250
   End
   Begin VB.CommandButton Command10 
      Caption         =   "<- Un Set Style"
      Height          =   540
      Left            =   3930
      TabIndex        =   14
      Top             =   6090
      Width           =   1005
   End
   Begin VB.CommandButton Command6 
      Caption         =   "<- Set Style"
      Height          =   540
      Left            =   2610
      TabIndex        =   13
      Top             =   6060
      Width           =   1035
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   135
      TabIndex        =   12
      Top             =   5940
      Width           =   2340
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<--- Attach"
      Height          =   495
      Left            =   2010
      TabIndex        =   10
      Top             =   5040
      Width           =   1635
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Visible on/off"
      Height          =   495
      Left            =   3930
      TabIndex        =   9
      Top             =   5055
      Width           =   1035
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Enable on/off"
      Height          =   495
      Left            =   3930
      TabIndex        =   8
      Top             =   7170
      Width           =   1035
   End
   Begin VB.CommandButton Command7 
      Caption         =   "<--- Scroll Bar Style"
      Height          =   495
      Left            =   2010
      TabIndex        =   6
      Top             =   7200
      Width           =   1635
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   150
      TabIndex        =   5
      Top             =   7050
      Width           =   1275
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   135
      TabIndex        =   4
      Top             =   4860
      Width           =   1275
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Size-"
      Height          =   675
      Left            =   3600
      TabIndex        =   3
      Top             =   3900
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Size+"
      Height          =   675
      Left            =   2520
      TabIndex        =   2
      Top             =   3900
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Detach"
      Height          =   675
      Left            =   1560
      TabIndex        =   1
      Top             =   3900
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Attach frmSplash"
      Height          =   675
      Left            =   255
      TabIndex        =   0
      Top             =   3900
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4800
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label2 
      Caption         =   "pHritz"
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "DanFratila@yahoo.com"
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   960
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Me.SubForm1.AttachForm frmSplash
End Sub


Private Sub Command10_Click()
  Select Case Me.List2.Text
    Case "None"
    Case "Moveable"

      Me.SubForm1.Moveable = False
    Case "Sizeable"
      Me.SubForm1.Sizeable = False
    Case "Moveable and Sizeable"
      Me.SubForm1.Moveable = False
      Me.SubForm1.Sizeable = False
  End Select

End Sub

Private Sub Command11_Click()
  Me.SubForm1.Caption = Me.Text2
End Sub

Private Sub Command2_Click()
  Me.SubForm1.DetachForm
End Sub

Private Sub Command3_Click()
  Me.SubForm1.Width = Me.SubForm1.Width * 1.1
End Sub

Private Sub Command4_Click()
  If Me.SubForm1.Width > 700 Then
    Me.SubForm1.Width = Me.SubForm1.Width * 0.9
  End If
End Sub

Private Sub Command5_Click()
Dim FRM As Form
Me.SubForm1.DetachForm

  If Me.List1.Text <> "" Then
    Select Case Me.List1.Text
        Case "frmSplash"
          Set FRM = New frmSplash
          Me.SubForm1.AttachForm FRM
        Case "frmAbout"
          Set FRM = New frmAbout
          Me.SubForm1.AttachForm FRM

        Case "frmForm2"
          Set FRM = New frmForm2
          Me.SubForm1.AttachForm FRM
        Case "this Form"
          Set FRM = New Form1
            Me.SubForm1.AttachForm FRM
        Case "frmTip"
          Set FRM = New frmTip
            Me.SubForm1.AttachForm FRM
    End Select


  End If
Set FRM = Nothing
End Sub


Private Sub Command6_Click()
  Select Case Me.List2.Text
    Case "None"
      Me.SubForm1.Moveable = False
      Me.SubForm1.Sizeable = False
    Case "Moveable"
      Me.SubForm1.Sizeable = False
      Me.SubForm1.Moveable = True
    Case "Sizeable"
      Me.SubForm1.Moveable = False
      Me.SubForm1.Sizeable = True
    Case "Moveable and Sizeable"
      Me.SubForm1.Moveable = True
      Me.SubForm1.Sizeable = True

  End Select
End Sub

Private Sub Command7_Click()
  Select Case Me.List3.Text
    Case "Regular"
      Me.SubForm1.ScrollBarStyle = efsRegular
    Case "Encarta"
      Me.SubForm1.ScrollBarStyle = efsEncarta
    Case "Flat"
      Me.SubForm1.ScrollBarStyle = efsFlat
  End Select
End Sub

Private Sub Command8_Click()
  Me.SubForm1.Visible = Not Me.SubForm1.Visible
End Sub

Private Sub Command9_Click()
  Me.SubForm1.Enabled = Not Me.SubForm1.Enabled
End Sub

Private Sub Form_Load()
Dim FRM As Form

  Set FRM = frmForm2
  Me.SubForm1.AttachForm frmForm2
  Set FRM = Nothing
  
  Set FRM = New frmAbout
  Me.SubForm2.AttachForm FRM
  Set FRM = Nothing
  
  With Me.List1
    .AddItem "frmSplash"
    .AddItem "frmAbout"
    .AddItem "frmForm2"
    .AddItem "this Form"
    .AddItem "frmTip"
  
    .Text = "frmForm2"
  End With
  
  With Me.List2
    .AddItem "None"
    .AddItem "Moveable"
    .AddItem "Sizeable"
    .AddItem "Moveable and Sizeable"
    .Text = "Moveable and Sizeable"
  End With
  
  With Me.List3
    .AddItem "Regular"
    .AddItem "Encarta"
    .AddItem "Flat"
    .Text = "Encarta"
  End With
  Me.Text2.Text = Me.SubForm1.Caption
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.SubForm2.DetachForm
  Me.SubForm1.DetachForm
End Sub
