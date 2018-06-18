VERSION 5.00
Begin VB.Form frmWP4 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "fx-95SG PLUS"
   ClientHeight    =   8715
   ClientLeft      =   495
   ClientTop       =   1050
   ClientWidth     =   4335
   Icon            =   "frmWP4.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmWP4.frx":45EAA
   ScaleHeight     =   581
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   Tag             =   "204"
   Begin VB.Timer adjFrm 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Label shiftS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   600
      TabIndex        =   4
      Top             =   1410
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label AlphaS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   810
      TabIndex        =   3
      Top             =   1410
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label stoS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "STO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1020
      TabIndex        =   2
      Top             =   1425
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label rclS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "RCL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1380
      TabIndex        =   1
      Top             =   1425
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image pButton 
      Height          =   345
      Index           =   49
      Left            =   1995
      Picture         =   "frmWP4.frx":C1446
      Tag             =   "200"
      Top             =   3150
      Width           =   345
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   48
      Left            =   3240
      Picture         =   "frmWP4.frx":C1B00
      Tag             =   "198"
      Top             =   7680
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   47
      Left            =   2535
      Picture         =   "frmWP4.frx":C2DDA
      Tag             =   "196"
      Top             =   7680
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   46
      Left            =   1830
      Picture         =   "frmWP4.frx":C40B4
      Tag             =   "194"
      Top             =   7680
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   45
      Left            =   1125
      Picture         =   "frmWP4.frx":C538E
      Tag             =   "192"
      Top             =   7680
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   44
      Left            =   420
      Picture         =   "frmWP4.frx":C6668
      Tag             =   "190"
      Top             =   7680
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   43
      Left            =   3240
      Picture         =   "frmWP4.frx":C7942
      Tag             =   "188"
      Top             =   7065
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   42
      Left            =   2535
      Picture         =   "frmWP4.frx":C8C1C
      Tag             =   "186"
      Top             =   7065
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   41
      Left            =   1830
      Picture         =   "frmWP4.frx":C9EF6
      Tag             =   "184"
      Top             =   7065
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   40
      Left            =   1125
      Picture         =   "frmWP4.frx":CB1D0
      Tag             =   "182"
      Top             =   7065
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   39
      Left            =   420
      Picture         =   "frmWP4.frx":CC4AA
      Tag             =   "180"
      Top             =   7065
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   38
      Left            =   3240
      Picture         =   "frmWP4.frx":CD784
      Tag             =   "178"
      Top             =   6450
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   37
      Left            =   2535
      Picture         =   "frmWP4.frx":CEA5E
      Tag             =   "176"
      Top             =   6450
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   36
      Left            =   1830
      Picture         =   "frmWP4.frx":CFD38
      Tag             =   "174"
      Top             =   6450
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   35
      Left            =   1125
      Picture         =   "frmWP4.frx":D1012
      Tag             =   "172"
      Top             =   6450
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   34
      Left            =   420
      Picture         =   "frmWP4.frx":D22EC
      Tag             =   "170"
      Top             =   6450
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   33
      Left            =   3240
      Picture         =   "frmWP4.frx":D35C6
      Tag             =   "168"
      Top             =   5835
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   32
      Left            =   2535
      Picture         =   "frmWP4.frx":D48A0
      Tag             =   "166"
      Top             =   5835
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   31
      Left            =   1830
      Picture         =   "frmWP4.frx":D5B7A
      Tag             =   "164"
      Top             =   5835
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   30
      Left            =   1125
      Picture         =   "frmWP4.frx":D6E54
      Tag             =   "162"
      Top             =   5835
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   510
      Index           =   29
      Left            =   420
      Picture         =   "frmWP4.frx":D812E
      Tag             =   "160"
      Top             =   5835
      Width           =   690
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   28
      Left            =   3345
      Picture         =   "frmWP4.frx":D9408
      Tag             =   "158"
      Top             =   5355
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   27
      Left            =   2760
      Picture         =   "frmWP4.frx":DA012
      Tag             =   "156"
      Top             =   5355
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   26
      Left            =   2175
      Picture         =   "frmWP4.frx":DAC1C
      Tag             =   "154"
      Top             =   5355
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   25
      Left            =   1590
      Picture         =   "frmWP4.frx":DB826
      Tag             =   "152"
      Top             =   5355
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   24
      Left            =   1005
      Picture         =   "frmWP4.frx":DC430
      Tag             =   "150"
      Top             =   5355
      Width           =   570
   End
   Begin VB.Label modal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "fx-95SG PLUS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   2820
      TabIndex        =   0
      Top             =   405
      Width           =   1125
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   23
      Left            =   420
      Picture         =   "frmWP4.frx":DD03A
      Tag             =   "148"
      Top             =   5355
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   22
      Left            =   3345
      Picture         =   "frmWP4.frx":DDC44
      Tag             =   "146"
      Top             =   4860
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   21
      Left            =   2760
      Picture         =   "frmWP4.frx":DE84E
      Tag             =   "144"
      Top             =   4860
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   20
      Left            =   2175
      Picture         =   "frmWP4.frx":DF458
      Tag             =   "142"
      Top             =   4860
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   19
      Left            =   1590
      Picture         =   "frmWP4.frx":E0062
      Tag             =   "140"
      Top             =   4860
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   18
      Left            =   1005
      Picture         =   "frmWP4.frx":E0C6C
      Tag             =   "138"
      Top             =   4860
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   17
      Left            =   420
      Picture         =   "frmWP4.frx":E1876
      Tag             =   "136"
      Top             =   4860
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   16
      Left            =   3345
      Picture         =   "frmWP4.frx":E2480
      Tag             =   "134"
      Top             =   4365
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   15
      Left            =   2760
      Picture         =   "frmWP4.frx":E308A
      Tag             =   "132"
      Top             =   4365
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   14
      Left            =   2175
      Picture         =   "frmWP4.frx":E3C94
      Tag             =   "130"
      Top             =   4365
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   13
      Left            =   1590
      Picture         =   "frmWP4.frx":E489E
      Tag             =   "128"
      Top             =   4365
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   12
      Left            =   1005
      Picture         =   "frmWP4.frx":E54A8
      Tag             =   "126"
      Top             =   4365
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   11
      Left            =   420
      Picture         =   "frmWP4.frx":E60B2
      Tag             =   "124"
      Top             =   4365
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   10
      Left            =   3375
      Picture         =   "frmWP4.frx":E6CBC
      Tag             =   "202"
      Top             =   3870
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   9
      Left            =   2790
      Picture         =   "frmWP4.frx":E78C8
      Tag             =   "120"
      Top             =   3870
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   8
      Left            =   975
      Picture         =   "frmWP4.frx":E84D2
      Tag             =   "118"
      Top             =   3870
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   390
      Index           =   7
      Left            =   390
      Picture         =   "frmWP4.frx":E90DC
      Tag             =   "116"
      Top             =   3870
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   495
      Index           =   6
      Left            =   3360
      Picture         =   "frmWP4.frx":E9CE6
      Tag             =   "114"
      Top             =   3105
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   495
      Index           =   5
      Left            =   2775
      Picture         =   "frmWP4.frx":EAC1C
      Tag             =   "112"
      Top             =   3180
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   495
      Index           =   4
      Left            =   1005
      Picture         =   "frmWP4.frx":EBB52
      Tag             =   "110"
      Top             =   3180
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   495
      Index           =   3
      Left            =   405
      Picture         =   "frmWP4.frx":ECA88
      Tag             =   "108"
      Top             =   3090
      Width           =   570
   End
   Begin VB.Image pButton 
      Height          =   345
      Index           =   2
      Left            =   2370
      Picture         =   "frmWP4.frx":ED9BE
      Tag             =   "106"
      Top             =   3450
      Width           =   345
   End
   Begin VB.Image pButton 
      Height          =   345
      Index           =   1
      Left            =   1620
      Picture         =   "frmWP4.frx":EE078
      Tag             =   "104"
      Top             =   3450
      Width           =   345
   End
   Begin VB.Image pButton 
      Height          =   345
      Index           =   0
      Left            =   1995
      Picture         =   "frmWP4.frx":EE732
      Tag             =   "102"
      Top             =   3750
      Width           =   345
   End
   Begin VB.Menu popMenu 
      Caption         =   "popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuOutPut 
         Caption         =   "Show/Hide OutPutWindow"
         Tag             =   "201"
      End
      Begin VB.Menu mnubar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
         Tag             =   "121"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Tag             =   "22"
      End
   End
End
Attribute VB_Name = "frmWP4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ax, bc As Long

Dim ShiftPower, AlphaPower, rclPower, stoPower As Boolean


Private Sub adjFrm_Timer()
    If keyCodeOutPut.Visible = True Then
        If keyCodeOutPut.Left <> Me.Left + Me.Width Or keyCodeOutPut.Top <> Me.Top Then
        keyCodeOutPut.Move Me.Left + Me.Width, Me.Top
        End If
    End If
End Sub

Private Sub Form_Load()
    Ax = 0
    bc = False
    Me.Picture = LoadResPicture(X + Val(Me.Tag), 0)
     '***************** Ðë¸´ÖÆ
    On Error Resume Next
    For Each Control In Me.Controls
        Control.Caption = LoadResString(Val(Control.Tag))
    Next Control
    
    
    'SHIFT ALPHA OPEN/CLOSE
    Let ShiftPower = False
    Let AlphaPower = False
    Let rclPower = False
    Let stoPower = False
    
    shiftS.Visible = ShiftPower
    AlphaS.Visible = AlphaPower
    stoS.Visible = stoPower
    rclS.Visible = rclPower
    
    'END
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu popmenu
    End If
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub
Private Sub modal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu popmenu
    Else
        If keyCodeOutPut.Visible = True Then
            keyCodeOutPut.btnShowList.SetFocus
        End If
    End If
End Sub
Private Sub mnuOutPut_Click()
    mnuOutPut.Checked = Not mnuOutPut.Checked

    
    
    
    If mnuOutPut.Checked Then
    mnuOutPut.Caption = LES(202)
    keyCodeOutPut.Show
    keyCodeOutPut.Move Me.Left + Me.Width, Me.Top
    Else
        mnuOutPut.Caption = LES(201)
        keyCodeOutPut.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'EndPlaySound
    Unload keyCodeOutPut
    frmAppMain.Enabled = True
End Sub


Private Sub pButton_DblClick(Index As Integer)
    pButton(Index).Picture = LoadResPicture(Ax + Val(pButton(Index).Tag) + 1, 0)
    bc = True
End Sub

Private Sub mnuAbout_Click()
    frmAppAbout.Show vbModal
End Sub



'*************************************************************************
'*************************************************************************
'*************************************************************************


Private Sub pButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu popmenu
        Exit Sub
    Else
        If keyCodeOutPut.Visible = True Then
            keyCodeOutPut.btnShowList.SetFocus
        End If
    End If
    
    '==================
    
    mnuOutPut.Checked = True
    
    
    mnuOutPut.Caption = LES(202)
    keyCodeOutPut.Show
    keyCodeOutPut.Move Me.Left + Me.Width, Me.Top
    
    
    '===================
    
    
    
    '===================
    
    
    
    If soundSwitch = True Then BeginPlaySound 126
    
    pButton(Index).Picture = LoadResPicture(Ax + Val(pButton(Index).Tag) + 1, 0)
    
    bc = True
    If keyCodeOutPut.Visible = False Then
        keyCodeOutPut.Show
    End If
    Select Case Index
    
    '============================Control Key============================
    
    Case 0
        keyCodeOutPut.btnShowList.AddItem "[¡ý]"
    Case 1
        keyCodeOutPut.btnShowList.AddItem "[¡û]"
    Case 2
        keyCodeOutPut.btnShowList.AddItem "[¡ú]"
    Case 3
        keyCodeOutPut.btnShowList.AddItem "[SHIFT]"
        
        'SHIFT OPEN/CLOSE
        If ShiftPower = False Then
        ShiftPower = True
        Else
        ShiftPower = False
        End If
        
        
        
        AlphaPower = False
        rclPower = False
        stoPower = False
        
        
        shiftS.Visible = ShiftPower
        AlphaS.Visible = AlphaPower
        
        stoS.Visible = stoPower
        rclS.Visible = rclPower
        
        
        Exit Sub

    Case 4
        keyCodeOutPut.btnShowList.AddItem "[ALPHA]"
        
        
        'ALPHA OPEN/CLOSE
        If AlphaPower = False Then
        AlphaPower = True
        Else
        AlphaPower = False
        End If
        
        ShiftPower = False
        rclPower = False
        stoPower = False
        
        shiftS.Visible = ShiftPower
        AlphaS.Visible = AlphaPower
        
        stoS.Visible = stoPower
        rclS.Visible = rclPower
        
        Exit Sub
        
    Case 5
        keyCodeOutPut.btnShowList.AddItem "[MODE]"
        
        If ShiftPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(SETUP)"
        Else
        End If
        
        
    Case 6
        keyCodeOutPut.btnShowList.AddItem "[ON]"
    
    '============================F Key============================
    
    Case 7
        keyCodeOutPut.btnShowList.AddItem "[Abs]"
        
        
        
    Case 8
        keyCodeOutPut.btnShowList.AddItem "[x^3]"
        
        
        If AlphaPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(:)"
        Else
        
        End If
    
    
    Case 9
        keyCodeOutPut.btnShowList.AddItem "[x^-1]"
        

        
    Case 10
        keyCodeOutPut.btnShowList.AddItem "[x!]"
    Case 11
        keyCodeOutPut.btnShowList.AddItem "[d/c]"
        
        If ShiftPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(ab/c)"
        Else
        End If
        
    Case 12
        keyCodeOutPut.btnShowList.AddItem "[¡Ì¡ö]"
        
        If ShiftPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(3¡Ì¡ö)"
        Else
        End If
        
    Case 13
        keyCodeOutPut.btnShowList.AddItem "[x^2]"
        
        
        
    Case 14
        keyCodeOutPut.btnShowList.AddItem "[x^¡ö]"
        
        If ShiftPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(¡ö¡Ì¡õ)"
        Else
        End If
        
    Case 15
        keyCodeOutPut.btnShowList.AddItem "[log]"
        
        If ShiftPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(10^¡ö)"
        Else
        End If
        
    Case 16
        keyCodeOutPut.btnShowList.AddItem "[ln]"
        
        If ShiftPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(e^¡ö)"
        Else
        End If
        
    Case 17
        keyCodeOutPut.btnShowList.AddItem "[(-)]"
        
        If AlphaPower = True Or rclPower = True Or stoPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(A)"
        Else
        End If
        
    Case 18
        keyCodeOutPut.btnShowList.AddItem "[dms]"
        
        
        If ShiftPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(¡û)"
        ElseIf AlphaPower = True Or rclPower = True Or stoPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(B)"
        Else
        
        End If
        
        
    Case 19
        keyCodeOutPut.btnShowList.AddItem "[hyp]"
        
        If AlphaPower = True Or rclPower = True Or stoPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(C)"
        Else
        End If
        
    Case 20
        keyCodeOutPut.btnShowList.AddItem "[sin]"
        
        
        If ShiftPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(sin-1)"
        ElseIf AlphaPower = True Or rclPower = True Or stoPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(D)"
        Else
        
        End If
        
        
    Case 21
        keyCodeOutPut.btnShowList.AddItem "[cos]"
        
        
        If ShiftPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(cos-1)"
        Else
        
        End If
        
        
    Case 22
        keyCodeOutPut.btnShowList.AddItem "[tan]"
        
        
        If ShiftPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(tan-1)"
        Else
        
        End If
        
        
    Case 23
        keyCodeOutPut.btnShowList.AddItem "[RCL]"
        
        If ShiftPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(STO)"
            
            If stoPower = False Then
                stoPower = True
                Else
                stoPower = False
                End If
                
                ShiftPower = False
                AlphaPower = False
                rclPower = False
                
                shiftS.Visible = ShiftPower
                AlphaS.Visible = AlphaPower
                
                
                stoS.Visible = stoPower
                rclS.Visible = rclPower
            
                Exit Sub
                
        Else
        End If
        
        'ALPHA OPEN/CLOSE
        If rclPower = False Then
        rclPower = True
        Else
        rclPower = False
        End If
        
        ShiftPower = False
        AlphaPower = False
        stoPower = False
        
        shiftS.Visible = ShiftPower
        AlphaS.Visible = AlphaPower
        
        
        stoS.Visible = stoPower
                rclS.Visible = rclPower
        
        Exit Sub
        
    Case 24
        keyCodeOutPut.btnShowList.AddItem "[ENG]"
        
        If ShiftPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(¡û)"
        Else
        End If
        
    Case 25
        keyCodeOutPut.btnShowList.AddItem "[(]"
        
        If ShiftPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(%)"
        Else
        End If
        
    Case 26
        keyCodeOutPut.btnShowList.AddItem "[)]"
        
        
        If ShiftPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(,)"
        ElseIf AlphaPower = True Or rclPower = True Or stoPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(X)"
        Else
        
        End If
        
        
    Case 27
        keyCodeOutPut.btnShowList.AddItem "[S-D]"
        
        
        If ShiftPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(ab/c-d/c)"
        ElseIf AlphaPower = True Or rclPower = True Or stoPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(Y)"
        Else
        
        End If
        
        
    Case 28
        keyCodeOutPut.btnShowList.AddItem "[M+]"
        
        
        If ShiftPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(M-)"
        ElseIf AlphaPower = True Or rclPower = True Or stoPower = True Then
            keyCodeOutPut.btnShowList.AddItem "(M)"
        Else
        
        End If
        
        
        
    '============================Number Key============================
    
         Case 29
             keyCodeOutPut.btnShowList.AddItem "[7]"
         Case 30
             keyCodeOutPut.btnShowList.AddItem "[8]"
         Case 31
             keyCodeOutPut.btnShowList.AddItem "[9]"
        
            If ShiftPower = True Then
                keyCodeOutPut.btnShowList.AddItem "(CLR)"
            Else
            End If
        
         Case 32
             keyCodeOutPut.btnShowList.AddItem "[DEL]"
        
            If ShiftPower = True Then
                keyCodeOutPut.btnShowList.AddItem "(INS)"
            Else
            End If
        
         Case 33
             keyCodeOutPut.btnShowList.AddItem "[AC]"
        
            If ShiftPower = True Then
                keyCodeOutPut.btnShowList.AddItem "(OFF)"
            Else
            End If
        
             
        
         
         Case 34
             keyCodeOutPut.btnShowList.AddItem "[4]"
         Case 35
             keyCodeOutPut.btnShowList.AddItem "[5]"
         Case 36
             keyCodeOutPut.btnShowList.AddItem "[6]"
         Case 37
             keyCodeOutPut.btnShowList.AddItem "[¡Á]"
        
            If ShiftPower = True Then
                keyCodeOutPut.btnShowList.AddItem "(nPr)"
            Else
            End If
        
         Case 38
             keyCodeOutPut.btnShowList.AddItem "[¡Â]"
        
            If ShiftPower = True Then
                keyCodeOutPut.btnShowList.AddItem "(nCr)"
            Else
            End If
        
             
        
         Case 39
             keyCodeOutPut.btnShowList.AddItem "[1]"
        
            If ShiftPower = True Then
                keyCodeOutPut.btnShowList.AddItem "(STAT)"
            Else
            End If
        
         Case 40
             keyCodeOutPut.btnShowList.AddItem "[2]"
         Case 41
             keyCodeOutPut.btnShowList.AddItem "[3]"
         Case 42
             keyCodeOutPut.btnShowList.AddItem "[+]"
        
            If ShiftPower = True Then
                keyCodeOutPut.btnShowList.AddItem "(Pol)"
            Else
            End If
        
         Case 43
             keyCodeOutPut.btnShowList.AddItem "[-]"
             
        
            If ShiftPower = True Then
                keyCodeOutPut.btnShowList.AddItem "(Rec)"
            Else
            End If
        
        
         
         Case 44
             keyCodeOutPut.btnShowList.AddItem "[0]"
        
            If ShiftPower = True Then
                keyCodeOutPut.btnShowList.AddItem "(Rnd)"
            Else
            End If
        
         Case 45
             keyCodeOutPut.btnShowList.AddItem "[.]"
             
              
            If ShiftPower = True Then
                keyCodeOutPut.btnShowList.AddItem "(Ran#)"
            ElseIf AlphaPower = True Then
                keyCodeOutPut.btnShowList.AddItem "(RanInt)"
            Else
            
            End If
        
         Case 46
             keyCodeOutPut.btnShowList.AddItem "[EXP]"
             
            If ShiftPower = True Then
                keyCodeOutPut.btnShowList.AddItem "(¦Ð)"
            ElseIf AlphaPower = True Then
                keyCodeOutPut.btnShowList.AddItem "(e)"
            Else
            
            End If
            
         Case 47
             keyCodeOutPut.btnShowList.AddItem "[Ans]"
        
            If ShiftPower = True Then
                keyCodeOutPut.btnShowList.AddItem "(DRG¡ú)"
            Else
            End If
        
         Case 48
             keyCodeOutPut.btnShowList.AddItem "[=]"
             
             
             
        'Up Key
        
        Case 49
             keyCodeOutPut.btnShowList.AddItem "[¡ü]"
    End Select
    
    
    
    
    '===================
    
    
    
        
    'SHIFT ALPHA OPEN/CLOSE
    Let ShiftPower = False
    Let AlphaPower = False
    Let rclPower = False
    Let stoPower = False
    
    
    shiftS.Visible = ShiftPower
    AlphaS.Visible = AlphaPower
    stoS.Visible = stoPower
    rclS.Visible = rclPower
    
    '===================

End Sub


'¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü
'¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü
'¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü¡ü

Private Sub pButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        pButton(Index).Picture = LoadResPicture(Ax + Val(pButton(Index).Tag) + 1, 0)
    Else
        pButton(Index).Picture = LoadResPicture(Ax + Val(pButton(Index).Tag), 0)
    End If
End Sub

Private Sub pButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bc = True Then
    For i = pButton.LBound To pButton.UBound
        pButton(i).Picture = LoadResPicture(Ax + Val(pButton(i).Tag), 0)
    Next i
    bc = False
    End If
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bc = True Then
    For i = pButton.LBound To pButton.UBound
        pButton(i).Picture = LoadResPicture(Ax + Val(pButton(i).Tag), 0)
    Next i
    bc = False
    End If
End Sub

