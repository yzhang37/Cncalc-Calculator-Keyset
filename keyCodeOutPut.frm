VERSION 5.00
Begin VB.Form keyCodeOutPut 
   BorderStyle     =   3  '固定对话框模式
   Caption         =   "OutPut"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   ControlBox      =   0   'False
   Icon            =   "keyCodeOutPut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   Tag             =   "25"
   Begin VB.ListBox btnShowList 
      Appearance      =   0  '平面
      DragIcon        =   "keyCodeOutPut.frx":000C
      BeginProperty Font 
         Name            =   "Helvetica"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4005
      IntegralHeight  =   0   'False
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      Width           =   4680
   End
   Begin VB.Menu popmenu 
      Caption         =   "popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnucopy 
         Caption         =   "CopyImage"
         Shortcut        =   ^A
         Tag             =   "132"
      End
      Begin VB.Menu bar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddItem 
         Caption         =   "AddItem"
         Tag             =   "29"
         Begin VB.Menu mnuAllKey 
            Caption         =   "AllKey"
            Tag             =   "35"
            Begin VB.Menu mnukey 
               Caption         =   "[SHIFT]"
               Index           =   0
            End
            Begin VB.Menu mnukey 
               Caption         =   "[SECONDE]"
               Index           =   1
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[ALPHA]"
               Index           =   2
            End
            Begin VB.Menu mnukey 
               Caption         =   "[MODE]"
               Index           =   3
            End
            Begin VB.Menu mnukey 
               Caption         =   "(SETUP)"
               Index           =   4
            End
            Begin VB.Menu mnukey 
               Caption         =   "(CONFIG)"
               Index           =   5
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[ON]"
               Index           =   6
            End
            Begin VB.Menu mnukey 
               Caption         =   "[←]"
               Index           =   7
            End
            Begin VB.Menu mnukey 
               Caption         =   "[↑]"
               Index           =   8
            End
            Begin VB.Menu mnukey 
               Caption         =   "[→]"
               Index           =   9
            End
            Begin VB.Menu mnukey 
               Caption         =   "[↓]"
               Index           =   10
            End
            Begin VB.Menu mnukey 
               Caption         =   "[CALC]"
               Index           =   11
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(SOLVE)"
               Index           =   12
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(=)"
               Index           =   13
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[Abs]"
               Index           =   14
            End
            Begin VB.Menu mnukey 
               Caption         =   "[%]"
               Index           =   15
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(Abs)"
               Index           =   16
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[x^2]"
               Index           =   17
            End
            Begin VB.Menu mnukey 
               Caption         =   "(√■)"
               Index           =   18
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[∫dx]"
               Index           =   19
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(d/dx)"
               Index           =   20
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[x^3]"
               Index           =   21
            End
            Begin VB.Menu mnukey 
               Caption         =   "(:)"
               Index           =   22
            End
            Begin VB.Menu mnukey 
               Caption         =   "[x^-1]"
               Index           =   23
            End
            Begin VB.Menu mnukey 
               Caption         =   "(x!)"
               Index           =   24
            End
            Begin VB.Menu mnukey 
               Caption         =   "(3√■)"
               Index           =   25
            End
            Begin VB.Menu mnukey 
               Caption         =   "[log■□]"
               Index           =   26
            End
            Begin VB.Menu mnukey 
               Caption         =   "(Σ)"
               Index           =   27
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[x^■]"
               Index           =   28
            End
            Begin VB.Menu mnukey 
               Caption         =   "(■√□)"
               Index           =   29
            End
            Begin VB.Menu mnukey 
               Caption         =   "[d/c]"
               Index           =   30
            End
            Begin VB.Menu mnukey 
               Caption         =   "(ab/c)"
               Index           =   31
            End
            Begin VB.Menu mnukey 
               Caption         =   "[√■]"
               Index           =   32
            End
            Begin VB.Menu mnukey 
               Caption         =   "(x^3)"
               Index           =   33
            End
            Begin VB.Menu mnukey 
               Caption         =   "[log]"
               Index           =   34
            End
            Begin VB.Menu mnukey 
               Caption         =   "(10^■)"
               Index           =   35
            End
            Begin VB.Menu mnukey 
               Caption         =   "[ln]"
               Index           =   36
            End
            Begin VB.Menu mnukey 
               Caption         =   "(e^■)"
               Index           =   37
            End
            Begin VB.Menu mnukey 
               Caption         =   "[DEC]"
               Index           =   38
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[HEX]"
               Index           =   39
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[BIN]"
               Index           =   40
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[OCT]"
               Index           =   41
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[Y]"
               Index           =   42
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[X]"
               Index           =   43
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[(-)]"
               Index           =   44
            End
            Begin VB.Menu mnukey 
               Caption         =   "(A)"
               Index           =   45
            End
            Begin VB.Menu mnukey 
               Caption         =   "(∠)"
               Index           =   46
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[dms]"
               Index           =   47
            End
            Begin VB.Menu mnukey 
               Caption         =   "(FACT)"
               Index           =   48
            End
            Begin VB.Menu mnukey 
               Caption         =   "(←)"
               Index           =   49
            End
            Begin VB.Menu mnukey 
               Caption         =   "(B)"
               Index           =   50
            End
            Begin VB.Menu mnukey 
               Caption         =   "[÷R]"
               Index           =   51
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(÷R)"
               Index           =   52
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[hyp]"
               Index           =   53
            End
            Begin VB.Menu mnukey 
               Caption         =   "(C)"
               Index           =   54
            End
            Begin VB.Menu mnukey 
               Caption         =   "[sin]"
               Index           =   55
            End
            Begin VB.Menu mnukey 
               Caption         =   "[cos]"
               Index           =   56
            End
            Begin VB.Menu mnukey 
               Caption         =   "[tan]"
               Index           =   57
            End
            Begin VB.Menu mnukey 
               Caption         =   "(sin-1)"
               Index           =   58
            End
            Begin VB.Menu mnukey 
               Caption         =   "(cos-1)"
               Index           =   59
            End
            Begin VB.Menu mnukey 
               Caption         =   "(tan-1)"
               Index           =   60
            End
            Begin VB.Menu mnukey 
               Caption         =   "(Asn)"
               Index           =   61
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(Acs)"
               Index           =   62
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(Atn)"
               Index           =   63
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(A)"
               Index           =   64
            End
            Begin VB.Menu mnukey 
               Caption         =   "(B)"
               Index           =   65
            End
            Begin VB.Menu mnukey 
               Caption         =   "(C)"
               Index           =   66
            End
            Begin VB.Menu mnukey 
               Caption         =   "(D)"
               Index           =   67
            End
            Begin VB.Menu mnukey 
               Caption         =   "(E)"
               Index           =   68
            End
            Begin VB.Menu mnukey 
               Caption         =   "(F)"
               Index           =   69
            End
            Begin VB.Menu mnukey 
               Caption         =   "[D]"
               Enabled         =   0   'False
               Index           =   70
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[E]"
               Enabled         =   0   'False
               Index           =   71
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[F]"
               Enabled         =   0   'False
               Index           =   72
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[RCL]"
               Index           =   73
            End
            Begin VB.Menu mnukey 
               Caption         =   "(STO)"
               Index           =   74
            End
            Begin VB.Menu mnukey 
               Caption         =   "[ENG]"
               Index           =   75
            End
            Begin VB.Menu mnukey 
               Caption         =   "(i)"
               Index           =   76
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[(]"
               Index           =   77
            End
            Begin VB.Menu mnukey 
               Caption         =   "[)]"
               Index           =   78
            End
            Begin VB.Menu mnukey 
               Caption         =   "(%)"
               Index           =   79
            End
            Begin VB.Menu mnukey 
               Caption         =   "(,)"
               Index           =   80
            End
            Begin VB.Menu mnukey 
               Caption         =   "[S-D]"
               Index           =   81
            End
            Begin VB.Menu mnukey 
               Caption         =   "(ab/c-d/c)"
               Index           =   82
            End
            Begin VB.Menu mnukey 
               Caption         =   "[M+]"
               Index           =   83
            End
            Begin VB.Menu mnukey 
               Caption         =   "(M-)"
               Index           =   84
            End
            Begin VB.Menu mnukey 
               Caption         =   "(X)"
               Index           =   85
            End
            Begin VB.Menu mnukey 
               Caption         =   "(Y)"
               Index           =   86
            End
            Begin VB.Menu mnukey 
               Caption         =   "(M)"
               Index           =   87
            End
            Begin VB.Menu mnukey 
               Caption         =   "[#]"
               Index           =   88
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(a×10^n)"
               Index           =   89
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(￣)"
               Index           =   90
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(*)"
               Index           =   91
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "((■))"
               Index           =   92
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[Simp]"
               Index           =   93
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[DEL]"
               Index           =   94
            End
            Begin VB.Menu mnukey 
               Caption         =   "[EFF]"
               Index           =   95
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[AC]"
               Index           =   96
            End
            Begin VB.Menu mnukey 
               Caption         =   "[+]"
               Index           =   97
            End
            Begin VB.Menu mnukey 
               Caption         =   "[-]"
               Index           =   98
            End
            Begin VB.Menu mnukey 
               Caption         =   "[×]"
               Index           =   99
            End
            Begin VB.Menu mnukey 
               Caption         =   "[÷]"
               Index           =   100
            End
            Begin VB.Menu mnukey 
               Caption         =   "[=]"
               Index           =   101
            End
            Begin VB.Menu mnukey 
               Caption         =   "[EXE]"
               Index           =   102
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[0]"
               Index           =   103
            End
            Begin VB.Menu mnukey 
               Caption         =   "[1]"
               Index           =   104
            End
            Begin VB.Menu mnukey 
               Caption         =   "[2]"
               Index           =   105
            End
            Begin VB.Menu mnukey 
               Caption         =   "[3]"
               Index           =   106
            End
            Begin VB.Menu mnukey 
               Caption         =   "[4]"
               Index           =   107
            End
            Begin VB.Menu mnukey 
               Caption         =   "[5]"
               Index           =   108
            End
            Begin VB.Menu mnukey 
               Caption         =   "[6]"
               Index           =   109
            End
            Begin VB.Menu mnukey 
               Caption         =   "[7]"
               Index           =   110
            End
            Begin VB.Menu mnukey 
               Caption         =   "[8]"
               Index           =   111
            End
            Begin VB.Menu mnukey 
               Caption         =   "[9]"
               Index           =   112
            End
            Begin VB.Menu mnukey 
               Caption         =   "[7F]"
               Index           =   113
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "[.]"
               Index           =   114
            End
            Begin VB.Menu mnukey 
               Caption         =   "[EXP]"
               Index           =   115
            End
            Begin VB.Menu mnukey 
               Caption         =   "[Ans]"
               Index           =   116
            End
            Begin VB.Menu mnukey 
               Caption         =   "(CONST)"
               Index           =   117
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(CONV)"
               Index           =   118
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(CLR)"
               Index           =   119
            End
            Begin VB.Menu mnukey 
               Caption         =   "(SUPPR)"
               Index           =   120
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(INS)"
               Index           =   121
            End
            Begin VB.Menu mnukey 
               Caption         =   "(OFF)"
               Index           =   122
            End
            Begin VB.Menu mnukey 
               Caption         =   "(MATRIX)"
               Index           =   123
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(VECTOR)"
               Index           =   124
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(nPr)"
               Index           =   125
            End
            Begin VB.Menu mnukey 
               Caption         =   "(nCr)"
               Index           =   126
            End
            Begin VB.Menu mnukey 
               Caption         =   "(STAT)"
               Index           =   127
            End
            Begin VB.Menu mnukey 
               Caption         =   "(CMPLX)"
               Index           =   128
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(VERIFY)"
               Index           =   129
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(BASE)"
               Index           =   130
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(Pol)"
               Index           =   131
            End
            Begin VB.Menu mnukey 
               Caption         =   "(Rec)"
               Index           =   132
            End
            Begin VB.Menu mnukey 
               Caption         =   "(Rnd)"
               Index           =   133
            End
            Begin VB.Menu mnukey 
               Caption         =   "(Ran#)"
               Index           =   134
            End
            Begin VB.Menu mnukey 
               Caption         =   "(RanInt)"
               Index           =   135
            End
            Begin VB.Menu mnukey 
               Caption         =   "(π)"
               Index           =   136
            End
            Begin VB.Menu mnukey 
               Caption         =   "(e)"
               Index           =   137
            End
            Begin VB.Menu mnukey 
               Caption         =   "(DRG→)"
               Index           =   138
            End
            Begin VB.Menu mnukey 
               Caption         =   "[x!]"
               Index           =   139
               Visible         =   0   'False
            End
            Begin VB.Menu mnukey 
               Caption         =   "(x^-1)"
               Index           =   140
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mnuAMode 
            Caption         =   "MODE"
            Begin VB.Menu mnuModeKey 
               Caption         =   "COMP"
               Index           =   0
            End
            Begin VB.Menu mnuModeKey 
               Caption         =   "CMPLX"
               Index           =   1
               Visible         =   0   'False
            End
            Begin VB.Menu mnuModeKey 
               Caption         =   "STAT"
               Index           =   2
            End
            Begin VB.Menu mnuModeKey 
               Caption         =   "BASE-N"
               Index           =   3
               Visible         =   0   'False
            End
            Begin VB.Menu mnuModeKey 
               Caption         =   "EQN"
               Index           =   4
               Visible         =   0   'False
            End
            Begin VB.Menu mnuModeKey 
               Caption         =   "INEQ"
               Index           =   5
               Visible         =   0   'False
            End
            Begin VB.Menu mnuModeKey 
               Caption         =   "RATIO"
               Index           =   6
               Visible         =   0   'False
            End
            Begin VB.Menu mnuModeKey 
               Caption         =   "PROP"
               Index           =   7
               Visible         =   0   'False
            End
            Begin VB.Menu mnuModeKey 
               Caption         =   "MATRIX"
               Index           =   8
               Visible         =   0   'False
            End
            Begin VB.Menu mnuModeKey 
               Caption         =   "VECTOR"
               Index           =   9
               Visible         =   0   'False
            End
            Begin VB.Menu mnuModeKey 
               Caption         =   "VRIFY"
               Index           =   10
               Visible         =   0   'False
            End
            Begin VB.Menu mnuModeKey 
               Caption         =   "TABLE"
               Index           =   11
            End
         End
         Begin VB.Menu mnuASetup 
            Caption         =   "SETUP"
            Begin VB.Menu mnuSetupKey 
               Caption         =   "MthIO"
               Index           =   0
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "LineIO"
               Index           =   1
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "MathO"
               Index           =   2
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "LineO"
               Index           =   3
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "Deg"
               Index           =   4
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "Rad"
               Index           =   5
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "Gra"
               Index           =   6
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "Fix"
               Index           =   7
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "Sci"
               Index           =   8
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "Norm"
               Index           =   9
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "ab/c"
               Index           =   10
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "d/c"
               Index           =   11
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "CMPLX"
               Index           =   12
               Visible         =   0   'False
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "a+bi"
               Index           =   13
               Visible         =   0   'False
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "r∠θ"
               Index           =   14
               Visible         =   0   'False
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "STAT"
               Index           =   15
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "ON"
               Index           =   16
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "OFF"
               Index           =   17
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "Disp"
               Index           =   18
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "Dot"
               Index           =   19
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "Comma"
               Index           =   20
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "Rdec"
               Index           =   21
            End
            Begin VB.Menu mnuSetupKey 
               Caption         =   "←CONT→"
               Index           =   22
            End
         End
         Begin VB.Menu mnuACmplx 
            Caption         =   "CMPLX"
            Visible         =   0   'False
            Begin VB.Menu mnuCmplxKey 
               Caption         =   "arg"
               Index           =   0
            End
            Begin VB.Menu mnuCmplxKey 
               Caption         =   "Conjg"
               Index           =   1
            End
            Begin VB.Menu mnuCmplxKey 
               Caption         =   "→r∠θ"
               Index           =   2
            End
            Begin VB.Menu mnuCmplxKey 
               Caption         =   "→a+bi"
               Index           =   3
            End
         End
         Begin VB.Menu mnuAStat 
            Caption         =   "STAT"
            Begin VB.Menu mnuStatKey 
               Caption         =   "1-VAR"
               Index           =   0
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "A+BX"
               Index           =   1
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "_+CX^2"
               Index           =   2
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "ln X"
               Index           =   3
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "e^x"
               Index           =   4
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "A*B^X"
               Index           =   5
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "A*X^B"
               Index           =   6
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "1/X"
               Index           =   7
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "Type"
               Index           =   8
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "Data"
               Index           =   9
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "Edit"
               Index           =   10
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "Sum"
               Index           =   11
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "Var"
               Index           =   12
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "MinMax"
               Index           =   13
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "Distr"
               Index           =   14
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "Reg"
               Index           =   15
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "Ins"
               Index           =   16
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "Del-A"
               Index           =   17
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "Σx^2"
               Index           =   18
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "Σx"
               Index           =   19
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "Σy^2"
               Index           =   20
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "Σy"
               Index           =   21
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "Σxy"
               Index           =   22
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "Σx^3"
               Index           =   23
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "Σx^2y"
               Index           =   24
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "Σx^4"
               Index           =   25
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "n"
               Index           =   26
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "￣-x"
               Index           =   27
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "xσn"
               Index           =   28
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "xσn-1"
               Index           =   29
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "￣-y"
               Index           =   30
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "yσn"
               Index           =   31
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "yσn-1"
               Index           =   32
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "minX"
               Index           =   33
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "maxX"
               Index           =   34
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "minY"
               Index           =   35
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "maxY"
               Index           =   36
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "P("
               Index           =   37
               Visible         =   0   'False
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "Q("
               Index           =   38
               Visible         =   0   'False
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "R("
               Index           =   39
               Visible         =   0   'False
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "→t"
               Index           =   40
               Visible         =   0   'False
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "A"
               Index           =   41
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "B"
               Index           =   42
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "r"
               Index           =   43
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "→x?"
               Index           =   44
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "→y?"
               Index           =   45
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "→x_1?"
               Index           =   46
            End
            Begin VB.Menu mnuStatKey 
               Caption         =   "→x_2?"
               Index           =   47
            End
         End
         Begin VB.Menu mnuABase 
            Caption         =   "BASE-N"
            Visible         =   0   'False
            Begin VB.Menu mnuBaseKey 
               Caption         =   "and"
               Index           =   0
            End
            Begin VB.Menu mnuBaseKey 
               Caption         =   "or"
               Index           =   1
            End
            Begin VB.Menu mnuBaseKey 
               Caption         =   "xor"
               Index           =   2
            End
            Begin VB.Menu mnuBaseKey 
               Caption         =   "xnor"
               Index           =   3
            End
            Begin VB.Menu mnuBaseKey 
               Caption         =   "Not"
               Index           =   4
            End
            Begin VB.Menu mnuBaseKey 
               Caption         =   "Neg"
               Index           =   5
            End
            Begin VB.Menu mnuBaseKey 
               Caption         =   "d"
               Index           =   6
            End
            Begin VB.Menu mnuBaseKey 
               Caption         =   "h"
               Index           =   7
            End
            Begin VB.Menu mnuBaseKey 
               Caption         =   "b"
               Index           =   8
            End
            Begin VB.Menu mnuBaseKey 
               Caption         =   "o"
               Index           =   9
            End
         End
         Begin VB.Menu mnuAEQN 
            Caption         =   "EQN"
            Visible         =   0   'False
            Begin VB.Menu mnuEqnKey 
               Caption         =   "a_nX+b_nY=c_n"
               Index           =   0
            End
            Begin VB.Menu mnuEqnKey 
               Caption         =   "a_nX+b_nY+c_nZ=d_n"
               Index           =   1
            End
            Begin VB.Menu mnuEqnKey 
               Caption         =   "aX^2+bX+c=0"
               Index           =   2
            End
            Begin VB.Menu mnuEqnKey 
               Caption         =   "aX^3+bX^2+cX+d=0"
               Index           =   3
            End
         End
         Begin VB.Menu mnuAINEQ 
            Caption         =   "INEQ"
            Visible         =   0   'False
            Begin VB.Menu mnuIneqKey 
               Caption         =   "aX^2+bX+c"
               Index           =   0
            End
            Begin VB.Menu mnuIneqKey 
               Caption         =   "aX^3+bX^2+cX+d"
               Index           =   1
            End
            Begin VB.Menu mnuIneqKey 
               Caption         =   "aX^2+bX+c<0"
               Index           =   2
            End
            Begin VB.Menu mnuIneqKey 
               Caption         =   "aX^2+bX+c>0"
               Index           =   3
            End
            Begin VB.Menu mnuIneqKey 
               Caption         =   "aX^2+bX+c<=0"
               Index           =   4
            End
            Begin VB.Menu mnuIneqKey 
               Caption         =   "aX^2+bX+c>=0"
               Index           =   5
            End
            Begin VB.Menu mnuIneqKey 
               Caption         =   "aX^3+bX^2+cX+d<0"
               Index           =   6
            End
            Begin VB.Menu mnuIneqKey 
               Caption         =   "aX^3+bX^2+cX+d>0"
               Index           =   7
            End
            Begin VB.Menu mnuIneqKey 
               Caption         =   "aX^3+bX^2+cX+d<=0"
               Index           =   8
            End
            Begin VB.Menu mnuIneqKey 
               Caption         =   "aX^3+bX^2+cX+d>=0"
               Index           =   9
            End
         End
         Begin VB.Menu mnuARatioPropGroup 
            Caption         =   "RATIO/PROP"
            Tag             =   "109"
            Visible         =   0   'False
            Begin VB.Menu mnuARatio 
               Caption         =   "RATIO"
               Begin VB.Menu mnuRatioKey 
                  Caption         =   "a:b=X:d"
                  Index           =   0
               End
               Begin VB.Menu mnuRatioKey 
                  Caption         =   "a:b=c:X"
                  Index           =   1
               End
            End
            Begin VB.Menu mnuAProp 
               Caption         =   "PROP"
               Begin VB.Menu mnuPropKey 
                  Caption         =   "a/b=X/d"
                  Index           =   0
               End
               Begin VB.Menu mnuPropKey 
                  Caption         =   "a/b=c/X"
                  Index           =   1
               End
            End
         End
         Begin VB.Menu mnuAMatVctGroup 
            Caption         =   "MATRIX/VECTOR"
            Tag             =   "110"
            Visible         =   0   'False
            Begin VB.Menu mnuAMatrix 
               Caption         =   "MATRIX"
               Begin VB.Menu mnuMatrixKey 
                  Caption         =   "Dim"
                  Index           =   0
               End
               Begin VB.Menu mnuMatrixKey 
                  Caption         =   "MatA"
                  Index           =   1
               End
               Begin VB.Menu mnuMatrixKey 
                  Caption         =   "MatB"
                  Index           =   2
               End
               Begin VB.Menu mnuMatrixKey 
                  Caption         =   "MatC"
                  Index           =   3
               End
               Begin VB.Menu mnuMatrixKey 
                  Caption         =   "3×3"
                  Index           =   4
               End
               Begin VB.Menu mnuMatrixKey 
                  Caption         =   "3×2"
                  Index           =   5
               End
               Begin VB.Menu mnuMatrixKey 
                  Caption         =   "3×1"
                  Index           =   6
               End
               Begin VB.Menu mnuMatrixKey 
                  Caption         =   "2×3"
                  Index           =   7
               End
               Begin VB.Menu mnuMatrixKey 
                  Caption         =   "2×2"
                  Index           =   8
               End
               Begin VB.Menu mnuMatrixKey 
                  Caption         =   "2×1"
                  Index           =   9
               End
               Begin VB.Menu mnuMatrixKey 
                  Caption         =   "1×3"
                  Index           =   10
               End
               Begin VB.Menu mnuMatrixKey 
                  Caption         =   "1×2"
                  Index           =   11
               End
               Begin VB.Menu mnuMatrixKey 
                  Caption         =   "1×1"
                  Index           =   12
               End
               Begin VB.Menu mnuMatrixKey 
                  Caption         =   "Data"
                  Index           =   13
               End
               Begin VB.Menu mnuMatrixKey 
                  Caption         =   "MatAns"
                  Index           =   14
               End
               Begin VB.Menu mnuMatrixKey 
                  Caption         =   "det"
                  Index           =   15
               End
               Begin VB.Menu mnuMatrixKey 
                  Caption         =   "Trn"
                  Index           =   16
               End
            End
            Begin VB.Menu mnuAVector 
               Caption         =   "VECTOR"
               Begin VB.Menu mnuVectorKey 
                  Caption         =   "Dim"
                  Index           =   0
               End
               Begin VB.Menu mnuVectorKey 
                  Caption         =   "VctA"
                  Index           =   1
               End
               Begin VB.Menu mnuVectorKey 
                  Caption         =   "VctB"
                  Index           =   2
               End
               Begin VB.Menu mnuVectorKey 
                  Caption         =   "VctC"
                  Index           =   3
               End
               Begin VB.Menu mnuVectorKey 
                  Caption         =   "3"
                  Index           =   4
               End
               Begin VB.Menu mnuVectorKey 
                  Caption         =   "2"
                  Index           =   5
               End
               Begin VB.Menu mnuVectorKey 
                  Caption         =   "VctAns"
                  Index           =   6
               End
               Begin VB.Menu mnuVectorKey 
                  Caption         =   "Data"
                  Index           =   7
               End
               Begin VB.Menu mnuVectorKey 
                  Caption         =   "Dot"
                  Index           =   8
               End
            End
         End
         Begin VB.Menu mnuAVrify 
            Caption         =   "VERIFY"
            Visible         =   0   'False
            Begin VB.Menu mnuVRIFYKey 
               Caption         =   "="
               Index           =   0
            End
            Begin VB.Menu mnuVRIFYKey 
               Caption         =   "≠"
               Index           =   1
            End
            Begin VB.Menu mnuVRIFYKey 
               Caption         =   ">"
               Index           =   2
            End
            Begin VB.Menu mnuVRIFYKey 
               Caption         =   "<"
               Index           =   3
            End
            Begin VB.Menu mnuVRIFYKey 
               Caption         =   "≥"
               Index           =   4
            End
            Begin VB.Menu mnuVRIFYKey 
               Caption         =   "≤"
               Index           =   5
            End
         End
         Begin VB.Menu mnuAClrSuppr 
            Caption         =   "CLR/SUPPR"
            Begin VB.Menu mnuClrSupprKey 
               Caption         =   "Setup"
               Index           =   0
            End
            Begin VB.Menu mnuClrSupprKey 
               Caption         =   "Memory"
               Index           =   1
            End
            Begin VB.Menu mnuClrSupprKey 
               Caption         =   "All"
               Index           =   2
            End
            Begin VB.Menu mnuClrSupprKey 
               Caption         =   "Yes"
               Index           =   3
            End
            Begin VB.Menu mnuClrSupprKey 
               Caption         =   "Cancel"
               Index           =   4
            End
            Begin VB.Menu mnuClrSupprKey 
               Caption         =   "Confg"
               Index           =   5
               Visible         =   0   'False
            End
            Begin VB.Menu mnuClrSupprKey 
               Caption         =   "Mém."
               Index           =   6
               Visible         =   0   'False
            End
            Begin VB.Menu mnuClrSupprKey 
               Caption         =   "Tout"
               Index           =   7
               Visible         =   0   'False
            End
            Begin VB.Menu mnuClrSupprKey 
               Caption         =   "Qui"
               Index           =   8
               Visible         =   0   'False
            End
            Begin VB.Menu mnuClrSupprKey 
               Caption         =   "Annuler"
               Index           =   9
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mnuADRG 
            Caption         =   "DRG→"
            Begin VB.Menu mnuDrgKey 
               Caption         =   "→Deg"
               Index           =   0
            End
            Begin VB.Menu mnuDrgKey 
               Caption         =   "→Rad"
               Index           =   1
            End
            Begin VB.Menu mnuDrgKey 
               Caption         =   "→Gra"
               Index           =   2
            End
         End
         Begin VB.Menu mnuCustomitem 
            Caption         =   "[CustomItem]"
            Tag             =   "411"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "DeleteItem"
         Shortcut        =   {DEL}
         Tag             =   "31"
      End
      Begin VB.Menu mnuRen 
         Caption         =   "RenameItem"
         Shortcut        =   {F2}
         Tag             =   "47"
         Visible         =   0   'False
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMoveUp 
         Caption         =   "moveUp"
         Tag             =   "45"
      End
      Begin VB.Menu mnuMoveDown 
         Caption         =   "moveDown"
         Tag             =   "46"
      End
   End
End
Attribute VB_Name = "keyCodeOutPut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctrlPower As Boolean

Private Sub btnShowList_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrE:
    Dim SelC As Integer
    'Caption = KeyCode
    Select Case KeyCode
        Case 17
        ctrlPower = True
        Exit Sub
        Case 46 '[DELETE]
            SelC = btnShowList.ListIndex
            
            btnShowList.RemoveItem btnShowList.ListIndex
            
            btnShowList.ListIndex = SelC
        Case 38
            If ctrlPower = True Then
            mnuMoveUp_Click
            btnShowList.Selected(btnShowList.ListIndex + 1) = True
            End If
        Case 40
            If ctrlPower = True Then
            mnuMoveDown_Click
            btnShowList.Selected(btnShowList.ListIndex - 1) = True
            End If
        Case 93
            PopupMenu popmenu
        Case 113
            If mnuRen.Visible = True Then
                mnuRen_Click
            End If
    End Select
ErrE:
    If Err = True Then
        Select Case Err.Number
        Case 380
            btnShowList.ListIndex = SelC - 1
        End Select
    End If
End Sub

Private Sub btnShowList_KeyUp(KeyCode As Integer, Shift As Integer)
    ctrlPower = False
End Sub

Private Sub btnShowList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu popmenu
    End If
End Sub




Private Sub Form_Load()
    On Error Resume Next
    For Each Control In Me.Controls
        Control.Caption = LoadResString(Val(Control.Tag))
    Next Control
    
    Caption = LES(Val(Me.Tag))
    ctrlPower = False
    
    Select Case Val(LES(61))
    Case 0
        showBf
        showCf
        showDf
        showEf
    Case 1
        showBf
        showCf
        showDf
    Case 2
        showBf
        showCf
    Case 3
        showBf
    Case 4
    End Select
End Sub

Private Sub mnuBaseKey_Click(Index As Integer)
    On Error GoTo E:
    
    If btnShowList.ListIndex <= -1 Then
    btnShowList.AddItem Replace("(%1)", "%1", mnuBaseKey(Index).Caption)
    btnShowList.Selected(btnShowList.ListCount - 1) = True
    Else
    btnShowList.AddItem Replace("(%1)", "%1", mnuBaseKey(Index).Caption), btnShowList.ListIndex + 1
    btnShowList.Selected(btnShowList.ListIndex + 1) = True
    End If
    
    '******************错误处理：复制一下内容******************
E:
    If Err Then
        MsgBox LES(106) & vbCrLf & Err.Number & vbCrLf & vbCrLf & LES(107) & "    " & Err.Description, vbExclamation, LES(104)
        Unload Me
    End If
End Sub

Private Sub mnuClrSupprKey_Click(Index As Integer)
    On Error GoTo E:
    
    If btnShowList.ListIndex <= -1 Then
    btnShowList.AddItem Replace("(%1)", "%1", mnuClrSupprKey(Index).Caption)
    btnShowList.Selected(btnShowList.ListCount - 1) = True
    Else
    btnShowList.AddItem Replace("(%1)", "%1", mnuClrSupprKey(Index).Caption), btnShowList.ListIndex + 1
    btnShowList.Selected(btnShowList.ListIndex + 1) = True
    End If
    
    '******************错误处理：复制一下内容******************
E:
    If Err Then
        MsgBox LES(106) & vbCrLf & Err.Number & vbCrLf & vbCrLf & LES(107) & "    " & Err.Description, vbExclamation, LES(104)
        Unload Me
    End If
End Sub

Private Sub mnuCmplxKey_Click(Index As Integer)
    On Error GoTo E:
    
    If btnShowList.ListIndex <= -1 Then
    btnShowList.AddItem Replace("(%1)", "%1", mnuCmplxKey(Index).Caption)
    btnShowList.Selected(btnShowList.ListCount - 1) = True
    Else
    btnShowList.AddItem Replace("(%1)", "%1", mnuCmplxKey(Index).Caption), btnShowList.ListIndex + 1
    btnShowList.Selected(btnShowList.ListIndex + 1) = True
    End If
    
    '******************错误处理：复制一下内容******************
E:
    If Err Then
        MsgBox LES(106) & vbCrLf & Err.Number & vbCrLf & vbCrLf & LES(107) & "    " & Err.Description, vbExclamation, LES(104)
        Unload Me
    End If
End Sub

Private Sub mnucopy_Click()
    Dim i As Integer, copyStr As String
    On Error GoTo E:
    copyStr = ""
    For i = 0 To btnShowList.ListCount - 1
        btnShowList.Selected(i) = True
        copyStr = copyStr & GetImageTxt(btnShowList.Text)
    Next
    Clipboard.Clear
    Clipboard.SetText copyStr
    
    '******************错误处理：复制一下内容******************
E:
    If Err Then
        MsgBox LES(106) & vbCrLf & Err.Number & vbCrLf & vbCrLf & LES(107) & "    " & Err.Description, vbExclamation, LES(104)
        Unload Me
    End If
End Sub

Private Sub mnuCustomitem_Click()
    
    Load CustomItem
    With CustomItem
    .Caption = LES(66)
    .ctlMth = AddNewItem
    .CustomItemTxt.Text = ""
    .Show vbModal, Me
    End With
        
    
End Sub

Private Sub mnuDelete_Click()
    On Error GoTo ErrE:
    Dim SelC As Integer
    SelC = btnShowList.ListIndex
            
            btnShowList.RemoveItem btnShowList.ListIndex
            
            btnShowList.ListIndex = SelC
ErrE:
    If Err = True Then
        Select Case Err.Number
        Case 380
            btnShowList.ListIndex = SelC - 1
        End Select
    End If
End Sub

Private Sub mnuDrgKey_Click(Index As Integer)
    On Error GoTo E:
    
    If btnShowList.ListIndex <= -1 Then
    btnShowList.AddItem Replace("(%1)", "%1", mnuDrgKey(Index).Caption)
    btnShowList.Selected(btnShowList.ListCount - 1) = True
    Else
    btnShowList.AddItem Replace("(%1)", "%1", mnuDrgKey(Index).Caption), btnShowList.ListIndex + 1
    btnShowList.Selected(btnShowList.ListIndex + 1) = True
    End If
    
    '******************错误处理：复制一下内容******************
E:
    If Err Then
        MsgBox LES(106) & vbCrLf & Err.Number & vbCrLf & vbCrLf & LES(107) & "    " & Err.Description, vbExclamation, LES(104)
        Unload Me
    End If
End Sub

Private Sub mnuEqnKey_Click(Index As Integer)
    On Error GoTo E:
    
    If btnShowList.ListIndex <= -1 Then
    btnShowList.AddItem Replace("(%1)", "%1", mnuEqnKey(Index).Caption)
    btnShowList.Selected(btnShowList.ListCount - 1) = True
    Else
    btnShowList.AddItem Replace("(%1)", "%1", mnuEqnKey(Index).Caption), btnShowList.ListIndex + 1
    btnShowList.Selected(btnShowList.ListIndex + 1) = True
    End If
    
    '******************错误处理：复制一下内容******************
E:
    If Err Then
        MsgBox LES(106) & vbCrLf & Err.Number & vbCrLf & vbCrLf & LES(107) & "    " & Err.Description, vbExclamation, LES(104)
        Unload Me
    End If
End Sub

Private Sub mnuIneqKey_Click(Index As Integer)
    On Error GoTo E:
    
    If btnShowList.ListIndex <= -1 Then
    btnShowList.AddItem Replace("(%1)", "%1", mnuIneqKey(Index).Caption)
    btnShowList.Selected(btnShowList.ListCount - 1) = True
    Else
    btnShowList.AddItem Replace("(%1)", "%1", mnuIneqKey(Index).Caption), btnShowList.ListIndex + 1
    btnShowList.Selected(btnShowList.ListIndex + 1) = True
    End If
    
    '******************错误处理：复制一下内容******************
E:
    If Err Then
        MsgBox LES(106) & vbCrLf & Err.Number & vbCrLf & vbCrLf & LES(107) & "    " & Err.Description, vbExclamation, LES(104)
        Unload Me
    End If
End Sub

Private Sub mnuKey_Click(Index As Integer)
    On Error GoTo E:
    
    If btnShowList.ListIndex <= -1 Then
    btnShowList.AddItem mnukey(Index).Caption
    btnShowList.Selected(btnShowList.ListCount - 1) = True
    Else
    btnShowList.AddItem mnukey(Index).Caption, btnShowList.ListIndex + 1
    btnShowList.Selected(btnShowList.ListIndex + 1) = True
    End If
    
    '******************错误处理：复制一下内容******************
E:
    If Err Then
        MsgBox LES(106) & vbCrLf & Err.Number & vbCrLf & vbCrLf & LES(107) & "    " & Err.Description, vbExclamation, LES(104)
        Unload Me
    End If
End Sub

Private Sub mnuMatrixKey_Click(Index As Integer)
    On Error GoTo E:
    
    If btnShowList.ListIndex <= -1 Then
    btnShowList.AddItem Replace("(%1)", "%1", mnuMatrixKey(Index).Caption)
    btnShowList.Selected(btnShowList.ListCount - 1) = True
    Else
    btnShowList.AddItem Replace("(%1)", "%1", mnuMatrixKey(Index).Caption), btnShowList.ListIndex + 1
    btnShowList.Selected(btnShowList.ListIndex + 1) = True
    End If
    
    '******************错误处理：复制一下内容******************
E:
    If Err Then
        MsgBox LES(106) & vbCrLf & Err.Number & vbCrLf & vbCrLf & LES(107) & "    " & Err.Description, vbExclamation, LES(104)
        Unload Me
    End If
End Sub

Private Sub mnuModeKey_Click(Index As Integer)
    On Error GoTo E:
    
    If btnShowList.ListIndex <= -1 Then
    btnShowList.AddItem Replace("(%1)", "%1", mnuModeKey(Index).Caption)
    btnShowList.Selected(btnShowList.ListCount - 1) = True
    Else
    btnShowList.AddItem Replace("(%1)", "%1", mnuModeKey(Index).Caption), btnShowList.ListIndex + 1
    btnShowList.Selected(btnShowList.ListIndex + 1) = True
    End If
    
    '******************错误处理：复制一下内容******************
E:
    If Err Then
        MsgBox LES(106) & vbCrLf & Err.Number & vbCrLf & vbCrLf & LES(107) & "    " & Err.Description, vbExclamation, LES(104)
        Unload Me
    End If
End Sub

Private Sub mnuMoveDown_Click()
    On Error Resume Next
  Dim nItem As Integer
  
  With btnShowList
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = .ListCount - 1 Then Exit Sub '不能将最后的项目向下移动
    '向下移动项目
    .AddItem .Text, nItem + 2
    '删除旧的项目
    .RemoveItem nItem
    '选择刚刚移动的项目
    .Selected(nItem + 1) = True
  End With

End Sub

Private Sub mnuMoveUp_Click()
    On Error Resume Next
    Dim nItem As Integer
  
    With btnShowList
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = 0 Then Exit Sub  '不能将第一个项目向上移动
    '向上移动项目
    .AddItem .Text, nItem - 1
    '删除旧项目
    .RemoveItem nItem + 1
    '选择刚刚移动的项目
    .Selected(nItem - 1) = True
  End With
End Sub

Private Sub mnuPropKey_Click(Index As Integer)
    On Error GoTo E:
    
    If btnShowList.ListIndex <= -1 Then
    btnShowList.AddItem Replace("(%1)", "%1", mnuPropKey(Index).Caption)
    btnShowList.Selected(btnShowList.ListCount - 1) = True
    Else
    btnShowList.AddItem Replace("(%1)", "%1", mnuPropKey(Index).Caption), btnShowList.ListIndex + 1
    btnShowList.Selected(btnShowList.ListIndex + 1) = True
    End If
    
    '******************错误处理：复制一下内容******************
E:
    If Err Then
        MsgBox LES(106) & vbCrLf & Err.Number & vbCrLf & vbCrLf & LES(107) & "    " & Err.Description, vbExclamation, LES(104)
        Unload Me
    End If
End Sub

Private Sub mnuRatioKey_Click(Index As Integer)
    On Error GoTo E:
    
    If btnShowList.ListIndex <= -1 Then
    btnShowList.AddItem Replace("(%1)", "%1", mnuRatioKey(Index).Caption)
    btnShowList.Selected(btnShowList.ListCount - 1) = True
    Else
    btnShowList.AddItem Replace("(%1)", "%1", mnuRatioKey(Index).Caption), btnShowList.ListIndex + 1
    btnShowList.Selected(btnShowList.ListIndex + 1) = True
    End If
    
    '******************错误处理：复制一下内容******************
E:
    If Err Then
        MsgBox LES(106) & vbCrLf & Err.Number & vbCrLf & vbCrLf & LES(107) & "    " & Err.Description, vbExclamation, LES(104)
        Unload Me
    End If
End Sub

Private Sub mnuRen_Click()
    Dim SelC As Integer
    SelC = btnShowList.ListIndex
    If SelC <= -1 Then
    MsgBox LES(112), vbExclamation, LES(69)
    Exit Sub
    End If
    Load CustomItem
    With CustomItem
    .Caption = LES(69)
    .ctlMth = RenameItem
    .CustomItemTxt.Text = btnShowList.Text
    .Show vbModal, Me
    End With
End Sub

Private Sub mnuSetupKey_Click(Index As Integer)
    On Error GoTo E:
    
    If btnShowList.ListIndex <= -1 Then
    btnShowList.AddItem Replace("(%1)", "%1", mnuSetupKey(Index).Caption)
    btnShowList.Selected(btnShowList.ListCount - 1) = True
    Else
    btnShowList.AddItem Replace("(%1)", "%1", mnuSetupKey(Index).Caption), btnShowList.ListIndex + 1
    btnShowList.Selected(btnShowList.ListIndex + 1) = True
    End If
    
    '******************错误处理：复制一下内容******************
E:
    If Err Then
        MsgBox LES(106) & vbCrLf & Err.Number & vbCrLf & vbCrLf & LES(107) & "    " & Err.Description, vbExclamation, LES(104)
        Unload Me
    End If
End Sub

Private Sub mnuStatKey_Click(Index As Integer)
    On Error GoTo E:
    
    If btnShowList.ListIndex <= -1 Then
    btnShowList.AddItem Replace("(%1)", "%1", mnuStatKey(Index).Caption)
    btnShowList.Selected(btnShowList.ListCount - 1) = True
    Else
    btnShowList.AddItem Replace("(%1)", "%1", mnuStatKey(Index).Caption), btnShowList.ListIndex + 1
    btnShowList.Selected(btnShowList.ListIndex + 1) = True
    End If
    
    '******************错误处理：复制一下内容******************
E:
    If Err Then
        MsgBox LES(106) & vbCrLf & Err.Number & vbCrLf & vbCrLf & LES(107) & "    " & Err.Description, vbExclamation, LES(104)
        Unload Me
    End If
End Sub

Private Sub mnuVectorKey_Click(Index As Integer)
    On Error GoTo E:
    
    If btnShowList.ListIndex <= -1 Then
    btnShowList.AddItem Replace("(%1)", "%1", mnuVectorKey(Index).Caption)
    btnShowList.Selected(btnShowList.ListCount - 1) = True
    Else
    btnShowList.AddItem Replace("(%1)", "%1", mnuVectorKey(Index).Caption), btnShowList.ListIndex + 1
    btnShowList.Selected(btnShowList.ListIndex + 1) = True
    End If
    
    '******************错误处理：复制一下内容******************
E:
    If Err Then
        MsgBox LES(106) & vbCrLf & Err.Number & vbCrLf & vbCrLf & LES(107) & "    " & Err.Description, vbExclamation, LES(104)
        Unload Me
    End If
End Sub

Private Sub mnuVRIFYKey_Click(Index As Integer)
    On Error GoTo E:
    
    If btnShowList.ListIndex <= -1 Then
    btnShowList.AddItem Replace("(%1)", "%1", mnuVRIFYKey(Index).Caption)
    btnShowList.Selected(btnShowList.ListCount - 1) = True
    Else
    btnShowList.AddItem Replace("(%1)", "%1", mnuVRIFYKey(Index).Caption), btnShowList.ListIndex + 1
    btnShowList.Selected(btnShowList.ListIndex + 1) = True
    End If
    
    '******************错误处理：复制一下内容******************
E:
    If Err Then
        MsgBox LES(106) & vbCrLf & Err.Number & vbCrLf & vbCrLf & LES(107) & "    " & Err.Description, vbExclamation, LES(104)
        Unload Me
    End If
End Sub

'型号设置
'****************************************************************************
'****************************************************************************
'****************************************************************************
'****************************************************************************


Private Function showBf()
    '******************mnukey
    mnukey(89).Visible = True
    mnukey(90).Visible = True
    mnukey(91).Visible = True
    mnukey(92).Visible = True
    mnukey(139).Visible = True
    
    '******************mnuModeKey
    mnuModeKey(4).Visible = True
    mnuModeKey(5).Visible = True
    mnuModeKey(6).Visible = True
    
    
    '******************mnuSetupKey
    mnuSetupKey(12).Visible = True
    mnuSetupKey(13).Visible = True
    mnuSetupKey(14).Visible = True
    
     '******************mnuAEQN
    mnuAEQN.Visible = True
    
     '******************mnuAINEQ
    mnuAINEQ.Visible = True
    
     '******************mnuARatioPropGroup
    mnuARatioPropGroup.Visible = True
    
End Function

Private Function hideBf()
    '******************mnukey
    mnukey(89).Visible = False
    mnukey(90).Visible = False
    mnukey(91).Visible = False
    mnukey(92).Visible = False
    mnukey(139).Visible = False
    
    '******************mnuModeKey
    mnuModeKey(4).Visible = False
    mnuModeKey(5).Visible = False
    mnuModeKey(6).Visible = False
    
    
    '******************mnuSetupKey
    mnuSetupKey(12).Visible = False
    mnuSetupKey(13).Visible = False
    mnuSetupKey(14).Visible = False
    
     '******************mnuAEQN
    mnuAEQN.Visible = False
    
     '******************mnuAINEQ
    mnuAINEQ.Visible = False
    
     '******************mnuARatioPropGroup
    mnuARatioPropGroup.Visible = False
End Function

Private Function showCf()
    '******************mnukey
    mnukey(1).Visible = True
    mnukey(5).Visible = True
    mnukey(16).Visible = True
    mnukey(18).Visible = True
    mnukey(42).Visible = True
    mnukey(43).Visible = True
    mnukey(51).Visible = True
    mnukey(61).Visible = True
    mnukey(62).Visible = True
    mnukey(63).Visible = True
    mnukey(95).Visible = True
    mnukey(102).Visible = True
    mnukey(113).Visible = True
    mnukey(120).Visible = True
    mnukey(129).Visible = True
    mnukey(140).Visible = True
    '******************mnuModeKey
    mnuModeKey(7).Visible = True
    mnuModeKey(10).Visible = True
   
    
    '******************mnuSetupKey
   mnuSetupKey(21).Visible = True
    
     '******************mnuAEQN
    
    
     '******************mnuAINEQ
    
    
     '******************mnuARatioPropGroup
    
    '******************mnuAVrify
    mnuAVrify.Visible = True
    
    
    '******************mnuClrSupprKey
    mnuClrSupprKey(5).Visible = True
    mnuClrSupprKey(6).Visible = True
    mnuClrSupprKey(7).Visible = True
    mnuClrSupprKey(8).Visible = True
    mnuClrSupprKey(9).Visible = True
End Function


Private Function hideCf()
    '******************mnukey
    mnukey(1).Visible = False
    mnukey(5).Visible = False
    mnukey(16).Visible = False
    mnukey(18).Visible = False
    mnukey(42).Visible = False
    mnukey(43).Visible = False
    mnukey(51).Visible = False
    mnukey(61).Visible = False
    mnukey(62).Visible = False
    mnukey(63).Visible = False
    mnukey(95).Visible = False
    mnukey(102).Visible = False
    mnukey(113).Visible = False
    mnukey(120).Visible = False
    mnukey(129).Visible = False
    mnukey(140).Visible = False
    '******************mnuModeKey
    mnuModeKey(7).Visible = False
    mnuModeKey(10).Visible = False
   
    
    '******************mnuSetupKey
    mnuSetupKey(21).Visible = False
    
     '******************mnuAEQN
    
    
     '******************mnuAINEQ
    
    
     '******************mnuARatioPropGroup
    
    '******************mnuAVrify
    mnuAVrify.Visible = False
    
    
    '******************mnuClrSupprKey
    mnuClrSupprKey(5).Visible = False
    mnuClrSupprKey(6).Visible = False
    mnuClrSupprKey(7).Visible = False
    mnuClrSupprKey(8).Visible = False
    mnuClrSupprKey(9).Visible = False
End Function

Private Function showDf()
    '******************mnukey
    mnukey(38).Visible = True
    mnukey(39).Visible = True
    mnukey(40).Visible = True
    mnukey(41).Visible = True
    mnukey(42).Visible = True
    mnukey(43).Visible = True
    mnukey(52).Visible = True
    mnukey(88).Visible = True
    mnukey(93).Visible = True
    mnukey(117).Visible = True
    mnukey(118).Visible = True
    mnukey(130).Visible = True
    
    '******************mnuModeKey
     mnuModeKey(3).Visible = True
      
    
    '******************mnuABase
   mnuABase.Visible = True
        
    
    
    mnuCustomitem.Visible = True
   
   mnuRen.Visible = True
End Function

Private Function hideDf()
    '******************mnukey
    mnukey(38).Visible = False
    mnukey(39).Visible = False
    mnukey(40).Visible = False
    mnukey(41).Visible = False
    mnukey(42).Visible = False
    mnukey(43).Visible = False
    mnukey(52).Visible = False
    mnukey(88).Visible = False
    mnukey(93).Visible = False
    mnukey(117).Visible = False
    mnukey(118).Visible = False
    mnukey(130).Visible = False
    
    '******************mnuModeKey
     mnuModeKey(3).Visible = False
      
    
    '******************mnuABase
   mnuABase.Visible = False
        
  mnuCustomitem.Visible = False
   
   mnuRen.Visible = False
End Function

Private Function showEf()
    '******************mnukey
    For i = mnukey.LBound To mnukey.UBound
    mnukey(i).Visible = True
    Next
    
    '******************mnuModeKey
    mnuModeKey(1).Visible = True
    mnuModeKey(8).Visible = True
    mnuModeKey(9).Visible = True
    
    
    
    
    '******************mnuACMPLX
   mnuACmplx.Visible = True
   
    '******************mnuStatKey
   mnuStatKey(37).Visible = True
   mnuStatKey(38).Visible = True
   mnuStatKey(39).Visible = True
   mnuStatKey(40).Visible = True
   
    '******************mnuAMatVctGroup
   mnuAMatVctGroup.Visible = True
   
   
   
   
End Function

Private Function hideEf()
    '******************mnukey
   For i = mnukey.LBound To mnukey.UBound
    mnukey(i).Visible = False
    Next
    
    '******************mnuModeKey
    mnuModeKey(1).Visible = False
    mnuModeKey(8).Visible = False
    mnuModeKey(9).Visible = False
    
    
    
    
    '******************mnuACMPLX
   mnuACmplx.Visible = False
   
    '******************mnuStatKey
   mnuStatKey(37).Visible = False
   mnuStatKey(38).Visible = False
   mnuStatKey(39).Visible = False
   mnuStatKey(40).Visible = False
   
    '******************mnuAMatVctGroup
   mnuAMatVctGroup.Visible = False
   
   
   
   
End Function


