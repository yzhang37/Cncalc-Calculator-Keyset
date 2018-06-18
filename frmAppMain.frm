VERSION 5.00
Begin VB.Form frmAppMain 
   AutoRedraw      =   -1  'True
   Caption         =   "fx-ES PLUS 字体获取器"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8190
   Icon            =   "frmAppMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   8190
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Align           =   4  '右端对其
      Appearance      =   0  '平面
      AutoRedraw      =   -1  'True
      BackColor       =   &H00502B10&
      BorderStyle     =   0  '无
      ForeColor       =   &H80000008&
      Height          =   5040
      Left            =   7590
      ScaleHeight     =   5040
      ScaleWidth      =   600
      TabIndex        =   10
      Top             =   1740
      Width           =   600
      Begin fxESFONT.jcbutton soundBtn 
         Height          =   690
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1217
         ButtonStyle     =   8
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
         Caption         =   ""
         Mode            =   1
         PictureNormal   =   "frmAppMain.frx":45EAA
         PictureHot      =   "frmAppMain.frx":46EFC
         PictureAlign    =   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin fxESFONT.jcbutton aboutBtn 
         Height          =   750
         Left            =   0
         TabIndex        =   15
         Top             =   675
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1323
         ButtonStyle     =   8
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
         Caption         =   ""
         PictureNormal   =   "frmAppMain.frx":47F4E
         PictureAlign    =   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
   End
   Begin VB.PictureBox Sec 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00502B10&
      BorderStyle     =   0  '无
      Height          =   5190
      Left            =   30
      Picture         =   "frmAppMain.frx":4A3A0
      ScaleHeight     =   5190
      ScaleWidth      =   8115
      TabIndex        =   1
      Top             =   1740
      Width           =   8115
      Begin fxESFONT.jcbutton openBtn 
         Height          =   840
         Index           =   0
         Left            =   255
         TabIndex        =   3
         Tag             =   "82"
         Top             =   555
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1482
         ButtonStyle     =   8
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
         Caption         =   "82ES PLUS"
         PictureNormal   =   "frmAppMain.frx":4A5E3
         PictureAlign    =   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin fxESFONT.jcbutton openBtn 
         Height          =   840
         Index           =   1
         Left            =   255
         TabIndex        =   4
         Tag             =   "83"
         Top             =   1425
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1482
         ButtonStyle     =   8
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
         Caption         =   "82ES PLUS"
         PictureNormal   =   "frmAppMain.frx":4CA35
         PictureAlign    =   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin fxESFONT.jcbutton openBtn 
         Height          =   840
         Index           =   2
         Left            =   255
         TabIndex        =   5
         Tag             =   "85"
         Top             =   2295
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1482
         ButtonStyle     =   8
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
         Caption         =   "82ES PLUS"
         PictureNormal   =   "frmAppMain.frx":4EE87
         PictureAlign    =   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin fxESFONT.jcbutton openBtn 
         Height          =   840
         Index           =   3
         Left            =   2655
         TabIndex        =   6
         Tag             =   "95"
         Top             =   555
         Visible         =   0   'False
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1482
         ButtonStyle     =   8
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
         Caption         =   "82ES PLUS"
         PictureNormal   =   "frmAppMain.frx":512D9
         PictureAlign    =   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin fxESFONT.jcbutton openBtn 
         Height          =   840
         Index           =   4
         Left            =   2655
         TabIndex        =   7
         Tag             =   "86"
         Top             =   1425
         Visible         =   0   'False
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1482
         ButtonStyle     =   8
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
         Caption         =   "82ES PLUS"
         PictureNormal   =   "frmAppMain.frx":5372B
         PictureAlign    =   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin fxESFONT.jcbutton openBtn 
         Height          =   840
         Index           =   5
         Left            =   2655
         TabIndex        =   8
         Tag             =   "500"
         Top             =   2295
         Visible         =   0   'False
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1482
         ButtonStyle     =   8
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
         Caption         =   "82ES PLUS"
         PictureNormal   =   "frmAppMain.frx":55B7D
         PictureAlign    =   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin fxESFONT.jcbutton openBtn 
         Height          =   840
         Index           =   6
         Left            =   2655
         TabIndex        =   9
         Tag             =   "92"
         Top             =   3165
         Visible         =   0   'False
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1482
         ButtonStyle     =   8
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
         Caption         =   "82ES PLUS"
         PictureNormal   =   "frmAppMain.frx":57FCF
         PictureAlign    =   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin fxESFONT.jcbutton openBtn 
         Height          =   840
         Index           =   7
         Left            =   255
         TabIndex        =   13
         Tag             =   "350"
         Top             =   3165
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1482
         ButtonStyle     =   8
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
         Caption         =   "82ES PLUS"
         PictureNormal   =   "frmAppMain.frx":5A421
         PictureAlign    =   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin fxESFONT.jcbutton openBtn 
         Height          =   840
         Index           =   8
         Left            =   255
         TabIndex        =   14
         Tag             =   "81"
         Top             =   4035
         Visible         =   0   'False
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1482
         ButtonStyle     =   8
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
         Caption         =   "82ES PLUS"
         PictureNormal   =   "frmAppMain.frx":5C873
         PictureAlign    =   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.Label lbl0 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "Please"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Index           =   0
         Left            =   285
         TabIndex        =   2
         Tag             =   "80"
         Top             =   195
         Width           =   510
      End
   End
   Begin VB.PictureBox title 
      Align           =   1  '顶端对其
      AutoRedraw      =   -1  'True
      BackColor       =   &H00D0A000&
      BorderStyle     =   0  '无
      Height          =   1740
      Left            =   0
      ScaleHeight     =   1740
      ScaleWidth      =   8190
      TabIndex        =   0
      Top             =   0
      Width           =   8190
      Begin VB.Label modalLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "fs-100SET"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   4635
         TabIndex        =   12
         Top             =   180
         Width           =   1515
      End
      Begin VB.Image logo 
         Height          =   900
         Left            =   270
         Picture         =   "frmAppMain.frx":5ECC5
         Top             =   210
         Width           =   2520
      End
   End
End
Attribute VB_Name = "frmAppMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aboutBtn_Click()
    frmAppAbout.Show vbModal
End Sub

Private Sub Form_Load()
    Dim i As Integer, s, t As String
    On Error Resume Next
    If App.PrevInstance = True Then
        MsgBox LES(33) & vbCrLf & LES(34), vbInformation, App.ProductName
        End
    End If
    
    For Each Control In Me.Controls
        Control.Caption = LoadResString(Val(Control.Tag))
    Next Control
    
    '=====================================================
        s = LES(60)
    For i = openBtn.LBound To openBtn.UBound
        
        t = Replace(s, "%2", LES(Val(openBtn(i).Tag)))
        openBtn(i).ToolTip = t
        
    Next
    soundSwitch = GetSetting(App.ProductName, "fxESPLUS", "SOUNDSet", 0)
    soundBtn.Value = soundSwitch
    
    
    Me.Move GetSetting(App.ProductName, "fxESPLUS Window Set", "Left", 0), GetSetting(App.ProductName, "fxESPLUS Window Set", "top", 0), _
                                        GetSetting(App.ProductName, "fxESPLUS Window Set", "width", 8200), GetSetting(App.ProductName, "fxESPLUS Window Set", "Height", 8000)
    Me.WindowState = GetSetting(App.ProductName, "fxESPLUS Window Set", "State", 0)
    
    
    
    modalLbl.Move ScaleWidth - 200 - modalLbl.Width, 200
    
    Select Case Val(LES(61))
    Case 0
        openBtn(8).Visible = True
        openBtn(3).Visible = True
        openBtn(5).Visible = True
        openBtn(6).Visible = True
        openBtn(4).Visible = True
        modalLbl.Caption = LES(271)
    Case 1
        openBtn(8).Visible = True
        openBtn(5).Visible = True
        openBtn(3).Visible = True
        openBtn(4).Visible = True
        modalLbl.Caption = LES(269)
    Case 2
        openBtn(8).Visible = True
        openBtn(3).Visible = True
        openBtn(4).Visible = True
        modalLbl.Caption = LES(267)
    Case 3
        openBtn(8).Visible = True
        openBtn(3).Visible = True
        modalLbl.Caption = LES(265)
    Case 4
        modalLbl.Caption = LES(233)
    End Select
    Show
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Sec.Move 0, title.Height, ScaleWidth, ScaleHeight - title.Height
    If Width < 8000 Then
        Width = 8000
    ElseIf Height < 7000 Then
        Height = 7000
    End If
    modalLbl.Move ScaleWidth - 200 - modalLbl.Width, 200
    title.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If soundSwitch = True Then
    SaveSetting App.ProductName, "fxESPLUS", "SOUNDSet", 1
    ElseIf soundSwitch = False Then
    SaveSetting App.ProductName, "fxESPLUS", "SOUNDSet", 0
    End If
    
    '位置
    
    SaveSetting App.ProductName, "fxESPLUS Window Set", "State", Me.WindowState
    Me.WindowState = 0
    Me.Hide
    SaveSetting App.ProductName, "fxESPLUS Window Set", "Left", Me.Left
    SaveSetting App.ProductName, "fxESPLUS Window Set", "Top", Me.Top
    SaveSetting App.ProductName, "fxESPLUS Window Set", "Width", Me.Width
    SaveSetting App.ProductName, "fxESPLUS Window Set", "Height", Me.Height
    End
End Sub

Private Sub logo_Click()
Shell "cmd.exe /c start " + Chr(34) + Chr(34) + " " + Chr(34) + "http://fxesms.5d6d.com" + Chr(34), vbHide
End Sub

Private Sub openBtn_Click(Index As Integer)
    Select Case Index
    Case 0
        frmWP1.Show
        frmWP1.Caption = LoadResString(12)
        frmWP1.modal = LES(12)
        frmWP1.Refresh
    Case 1
        frmWP2.Show
        frmWP2.Caption = LoadResString(12)
        frmWP2.modal = LES(12)
        frmWP2.Refresh
    Case 2
        frmWP3.Show
        frmWP3.Caption = LoadResString(13)
        frmWP3.modal = LES(13)
        frmWP3.Refresh
    Case 3
        frmWP4.Show
        frmWP4.Caption = LoadResString(16)
        frmWP4.modal = LES(16)
        frmWP4.Refresh
    Case 4
        frmWP5.Show
        frmWP5.Caption = LoadResString(14)
        frmWP5.modal = LES(14)
        frmWP5.Refresh
    Case 5
        frmWP6.Show
        frmWP6.Caption = LoadResString(18)
        frmWP6.modal = LES(18)
        frmWP6.Refresh
    Case 6
        frmWP7.Show
        frmWP7.Caption = LoadResString(15)
        'frmWP7.modal = LES(15)
        frmWP7.Refresh
    Case 7
        frmWP1.Show
        frmWP1.Caption = LoadResString(17)
        frmWP1.modal = LES(17)
        frmWP1.Refresh
    Case 8
        frmWP4.Show
        frmWP4.Caption = LoadResString(11)
        frmWP4.modal = LES(11)
        frmWP4.Refresh
    End Select
    Me.Enabled = False
End Sub


Private Sub soundbtn_Click()
    soundSwitch = soundBtn.Value
End Sub

