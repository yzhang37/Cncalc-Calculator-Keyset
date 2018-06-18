VERSION 5.00
Begin VB.Form CustomItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NEW/RENAME Item"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3840
   Icon            =   "CustomItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin fxESFONT.jcbutton btnSel 
      Default         =   -1  'True
      Height          =   330
      Index           =   0
      Left            =   1755
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "67"
      Top             =   720
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   582
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
   Begin VB.TextBox CustomItemTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   105
      MaxLength       =   30
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   345
      Width           =   3660
   End
   Begin fxESFONT.jcbutton btnSel 
      Cancel          =   -1  'True
      Height          =   330
      Index           =   1
      Left            =   2745
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "75"
      Top             =   720
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   582
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
      Caption         =   "Cancel"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Label namItemLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   180
      Left            =   135
      TabIndex        =   1
      Tag             =   "63"
      Top             =   105
      Width           =   360
   End
End
Attribute VB_Name = "CustomItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ctlMth As NameItem
Private Sub btnSel_Click(Index As Integer)
    On Error GoTo E:
    Dim SelC
    Select Case Index
    Case 0
    '========================================================
        Select Case ctlMth
        
        Case AddNewItem
        
            If Me.CustomItemTxt = "" Then
                    MsgBox LES(113), vbExclamation, LES(71)
                Exit Sub
            End If
            
            With keyCodeOutPut.btnShowList
            SelC = .ListIndex
            
            If SelC <= -1 Then
            .AddItem Me.CustomItemTxt
            .Selected(.ListCount - 1) = True
            
            Else
            .AddItem Me.CustomItemTxt, SelC + 1
            
            .Selected(SelC + 1) = True
            End If
            
            
            
            End With
            
            
        Case RenameItem
        
            With keyCodeOutPut.btnShowList
            SelC = .ListIndex
            
            .RemoveItem .ListIndex
            
            .AddItem Me.CustomItemTxt, SelC
            
            .Selected(SelC) = True
            
            End With
            
            
        End Select
        
    '========================================================
    End Select
    Unload Me
    
E:
    If Err Then
        MsgBox LES(106) & vbCrLf & Err.Number & vbCrLf & vbCrLf & LES(107) & "    " & Err.Description, vbExclamation, LES(104)
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    For Each Control In Me.Controls
        Control.Caption = LoadResString(Val(Control.Tag))
        Control.ToolTip = LoadResString(100 + Val(Control.Tag))
    Next Control
End Sub
