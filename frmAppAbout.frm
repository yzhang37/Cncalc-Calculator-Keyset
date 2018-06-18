VERSION 5.00
Begin VB.Form frmAppAbout 
   BorderStyle     =   3  '固定对话框模式
   Caption         =   "About"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   Icon            =   "frmAppAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin fxESFONT.jcbutton btnok 
      Default         =   -1  'True
      Height          =   375
      Left            =   4395
      TabIndex        =   1
      Tag             =   "67"
      Top             =   2955
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      ButtonStyle     =   8
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "OK"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin fxESFONT.jcbutton showIcon 
      Height          =   3960
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   6985
      ButtonStyle     =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Caption         =   ""
      PictureNormal   =   "frmAppAbout.frx":000C
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin fxESFONT.jcbutton btnSysInfo 
      Height          =   375
      Left            =   4395
      TabIndex        =   5
      Tag             =   "93"
      Top             =   3390
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      ButtonStyle     =   8
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "SysInfo"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "111"
      Height          =   180
      Left            =   2475
      TabIndex        =   4
      Top             =   1350
      Width           =   3345
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblversion 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "111"
      Height          =   180
      Left            =   2460
      TabIndex        =   3
      Top             =   930
      Width           =   3345
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblProd 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "ProductName"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2460
      TabIndex        =   2
      Top             =   240
      Width           =   3270
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAppAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnok_Click()
    Unload Me
End Sub

Private Sub btnSysInfo_Click()
    Call StartSysInfo
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'unload Me
End Sub

Private Sub Form_Load()
    Caption = LES(73)
    
    On Error Resume Next
    For Each Control In Me.Controls
        Control.Caption = LoadResString(Val(Control.Tag))
    Next Control
    
     lblversion.Left = lblProd.Left
    lblComment.Left = lblProd.Left
    lblProd = App.ProductName
    
    lblversion.Top = lblProd.Top + lblProd.Height + 200
    lblversion = LES(137) & ": " & App.Major & "." & App.Minor & "." & App.Revision & "("
    Select Case Val(LES(61))
    Case 0
        lblversion = lblversion & LES(271) & ")"
    Case 1
        lblversion = lblversion & LES(269) & ")"
    Case 2
        lblversion = lblversion & LES(267) & ")"
    Case 3
        lblversion = lblversion & LES(265) & ")"
    Case 4
        lblversion = lblversion & LES(233) & ")"
    End Select
    
    
    
    lblComment.Top = lblversion.Top + lblversion.Height + 200
    lblComment = LES(199) & vbCrLf & LES(125) & vbCrLf & LES(138) & vbCrLf & LES(176) & vbCrLf & LES(177)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Unload Me
End Sub

Private Sub Form_Resize()
    showIcon.Move 0, 0
End Sub

Private Sub jcbutton1_Click()
    
End Sub
