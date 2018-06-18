Attribute VB_Name = "copyPicture"

Public Function GetImageTxt(ImageTxt As String) As String
Dim t As String
Select Case UCase(ImageTxt)
    '******************CTRL KEYS******************
    Case "[SHIFT]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262500451bC5e.jpg")
    Case "[ALPHA]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507970jvZG.jpg")
    Case "[MODE]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501280TsII.jpg")
    Case "[SECONDE]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501321q7Vv.jpg")
    Case "[ON]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_126250129513Xp.jpg")
    
        
    '******************DR KEYS******************
        
    Case "[¡ü]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507975yW3y.jpg")
    Case "[¡û]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262502032El6m.jpg")
    Case "[¡ú]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262502034W6lj.jpg")
    Case "[¡ý]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262502030cxOM.jpg")
        
    '******************FUNCTION KEYS******************
    
    Case "[CALC]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_126250797078N7.jpg")
    Case "[ABS]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507969B3ZW.jpg")
    Case "[%]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_12625079136qD4.jpg")
    Case "[X^2]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501383ZFe1.jpg")
    Case "[X^3]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501380zRR6.jpg")
    Case "[X^-1]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501387RZ80.jpg")
    Case "(D/DX)"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501086Mk83.jpg")
    Case "(¡Ì¡ö)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/19/9696253_126660148755lx.png")
    Case "[¡ÒDX]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501290C1ac.jpg")
    Case "(3¡Ì¡ö)"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507900CYTt.jpg")
    Case "[log¡ö¡õ]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507975fANf.jpg")
    Case "(¦²)"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_12625079002q11.jpg")
    Case "[X!]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_12625013980pgn.jpg")
    Case "[X^¡ö]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501414QaKW.jpg")
    Case "(X^-1)"
        t = "([font=Times New Roman][i]x[/i][/font][sup]-1[/sup])"
    'SECOND
        
    Case "[D/C]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507974ix1M.jpg")
    Case "(AB/C)"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501425FAFi.jpg")
    Case "[¡Ì¡ö]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501335gOUh.jpg")
    Case "[Y]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501417bf0N.jpg")
    Case "[X]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_12625014114X4p.jpg")
    Case "(¡ö¡Ì¡õ)"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_12625078941nl2.jpg")
    Case "(X^3)"
        t = "([font=Times New Roman]x[/font][sup]3[/sup])"
    Case "(10^¡ö)"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262500319SEYS.jpg")
    Case "[LOG]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_12625005818BqZ.jpg")
    Case "[LN]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507975Mk9R.jpg")
    Case "(E^¡ö)"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507891eEWO.jpg")
    Case "(a¡Á10^n)"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501052zG69.jpg")
    Case "[DEC]"  'NEW
        t = SetURLtoImg("http://img1.5d6d.net/201002/17/9696253_1266432309cYP1.png")
    Case "[HEX]"  'NEW
        t = SetURLtoImg("http://img1.5d6d.net/201002/17/9696253_1266432310e068.png")
    Case "[BIN]"  'NEW
        t = SetURLtoImg("http://img1.5d6d.net/201002/17/9696253_126643230984cR.png")
    Case "[OCT]"  'NEW
        t = SetURLtoImg("http://img1.5d6d.net/201002/17/9696253_1266432310L10l.png")
    Case "((¡ö))"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501037f4cd.jpg")
    Case "(£þ)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/10/9696253_1265788842790u.png")
    Case "(*)"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501046rtxL.jpg")
        
    'THIRD
    Case "[(-)]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507913wf5w.jpg")
    Case "(¡Ï)"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_12625010804m51.jpg")
    Case "[DMS]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507972HKwc.jpg")
    Case "(¡û)"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_12625078934Pp0.jpg")
    Case "[¡ÂR]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_126250130502ED.jpg")
    Case "[HYP]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_12625079749Qpr.jpg")
    Case "[SIN]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501330hfFS.jpg")
    Case "[COS]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507971P6Ep.jpg")
    Case "[TAN]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501344rKk6.jpg")
    Case "(SIN-1)"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501066q9AI.jpg")
    Case "(COS-1)"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501062U217.jpg")
    Case "(TAN-1)"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501070iY57.jpg")
        
    
    'FOURTH
    Case "[RCL]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501309Pf84.jpg")
    Case "[ENG]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507973jYN1.jpg")
    Case "(I)"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_12625078923a3g.jpg")
    Case "[(]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507913bjWI.jpg")
    Case "[)]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507914CZz8.jpg")
    Case "(%)"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_12625010316ooQ.jpg")
    Case "(,)"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501041mK7i.jpg")
    Case "[S-D]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501319YYSh.jpg")
    Case "(AB/C-D/C)"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_12625010579It7.jpg")
    Case "[M+]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501278Cq9O.jpg")
    
    
    'NUMBER KEY
    
    'FIRST
    Case "[7]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_12625011482x63.jpg")
    Case "[8]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_12625011625YMr.jpg")
    Case "[9]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501273333n.jpg")
    Case "[SIMP]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_12625013251U00.jpg")
    Case "[EFF]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507972vLh1.jpg")
    Case "[DEL]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_12625012852sci.jpg")
    Case "[AC]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_12625079697H7i.jpg")
    Case "[7F]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501155wgi8.jpg")
        
    'SECOND
    Case "[4]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_12625011294gGJ.jpg")
    Case "[5]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501137RST5.jpg")
    Case "[6]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501142Et3z.jpg")
    Case "[¡Á]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501366NI1o.jpg")
    Case "[¡Â]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507971p294.jpg")
    'THIRD
    Case "[1]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501106JVw2.jpg")
    Case "[2]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501117c27L.jpg")
    Case "[3]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501122Q5mx.jpg")
    Case "[+]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507922ITuC.jpg")
    Case "[-]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507921QVq5.jpg")
    
    
    'FOURTH
    Case "[0]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501096yynF.jpg")
    Case "[.]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507915ccta.jpg")
    Case "[EXP]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507912DaHE.jpg")
    Case "[ANS]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507970Sb9d.jpg")
    Case UCase("(¦Ð)")
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507894UEVC.jpg")
    Case "(E)" 'NEW
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266484822V4T7.png")
    Case "(DRG¡ú)"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262501421XXG4.jpg")
    Case "[=]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_1262507927StVt.jpg")
    Case "[EXE]"
        t = SetURLtoImg("http://img1.5d6d.net/201001/3/9696253_12625079734653.jpg")
        
    '**************** CMPLX
    Case "(¡úA+BI)" 'NEW
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266483474q1vh.png")
    Case "(¡úR¡Ï¦¨)" 'NEW
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266483474hO0e.png")
    
    '****************STAT
    
    '******* TYPE
    Case "(_+CX^2)"
         t = "(_+CX[sup]2[/sup])"
    Case "(E^X)"
         t = "(e[sup]x[/sup])"
    Case "(A*B^X)"
         t = "(A¡¤B[sup]x[/sup])"
    Case "(A*X^B)"
         t = "(A¡¤[font=Times New Roman][i]X[/i][/font][sup]B[/sup])"
    
    
    
    
    '******* VAR
    
    'ALL NEW
    
    Case "(N)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_12664848267xRx.png")
    Case "(£þ-X)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266484822vA77.png")
    Case "(X¦²N)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266484822uz2F.png")
    Case "(X¦²N-1)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266484823jSZ6.png")
    Case "(£þ-Y)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266484823dh9Z.png")
    Case "(Y¦²N)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266484824y5d2.png")
    Case "(Y¦²N-1)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266484824xekv.png")
    
     '******* SUM *******
    
    'ALL NEW
    Case "(¦²X^2)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266483478BcZb.png")
    Case "(¦²X)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266483478m1Hf.png")
    Case "(¦²Y^2)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266483478gqIQ.png")
    Case "(¦²Y)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266483479mxtt.png")
    Case "(¦²XY)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_12664834797f7I.png")
    Case "(¦²X^3)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266483480ZsMD.png")
    Case "(¦²X^2Y)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266483480CEPk.png")
    Case "(¦²X^4)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266483554s58I.png")
    
     '******* DISTR *******
    
    'ALL NEW
    
    Case "(¡úT)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266485838MDRw.png")
    
    '******* REG *******
    
    'ALL NEW
    Case "(¡úX?)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266484825004R.png")
    Case "(¡úY?)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266484823dh9Z.png")
    Case "(¡úX_1?)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_12664848261H3Z.png")
    Case "(¡úX_2?)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_1266484826w8lH.png")
    Case "(R)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/18/9696253_12664848259l6A.png")
        
        
    '******* EQN *******
    
    Case "(A_NX+B_NY=C_N)"
        t = "(a[sub]n[/sub]X+b[sub]n[/sub]Y=c[sub]n[/sub])"
    Case "(A_NX+B_NY+C_NZ=D_N)"
        t = "(a[sub]n[/sub]X+b[sub]n[/sub]Y+c[sub]n[/sub]Z=d[sub]n[/sub])"
    Case "(AX^2+BX+C=0)"
        t = "(aX[sup]2[/sup]+bX+c=0)"
    Case "(AX^3+BX^2+CX+D=0)"
        t = "(aX[sup]3[/sup]+bX[sup]2[/sup]+cX+d=0)"
        
    '******* INEQ *******
    
    Case "(AX^2+BX+C)"
        t = "(aX[sup]2[/sup]+bX+c)"
    Case "(AX^3+BX^2+CX+D)"
        t = "(aX[sup]3[/sup]+bX[sup]2[/sup]+cX+d)"
    
    
    Case "(AX^2+BX+C<0)"
        t = "(aX[sup]2[/sup]+bX+c<0)"
    Case "(AX^2+BX+C>0)"
        t = "(aX[sup]2[/sup]+bX+c>0)"
    Case "(AX^2+BX+C<=0)"
        t = "(aX[sup]2[/sup]+bX+c¡Ü0)"
    Case "(AX^2+BX+C>=0)"
        t = "(aX[sup]2[/sup]+bX+c¡Ý0)"
    
    Case "(AX^3+BX^2+CX+D<0)"
        t = "(aX[sup]3[/sup]+bX[sup]2[/sup]+cX+d<0)"
    Case "(AX^3+BX^2+CX+D>0)"
        t = "(aX[sup]3[/sup]+bX[sup]2[/sup]+cX+d>0)"
    Case "(AX^3+BX^2+CX+D<=0)"
        t = "(aX[sup]3[/sup]+bX[sup]2[/sup]+cX+d¡Ü0)"
    Case "(AX^3+BX^2+CX+D>=0)"
        t = "(aX[sup]3[/sup]+bX[sup]2[/sup]+cX+d¡Ý0)"
    '**************** DRG
    Case "(¡úDEG)"
        t = "(¡ã)"
    Case "(¡úRAD)"
        t = "([sup]r[/sup])"
    Case "(¡úGRA)"
        t = "([sup]g[/sup])"
        
    '**************** SETUP
    Case "(¡ûCONT¡ú)"
        t = SetURLtoImg("http://img1.5d6d.net/201002/26/9696253_12672042966oQD.png")
        
    Case "(A+BI)"
        t = "([font=Times New Roman][i]a[/i]+[i]b[b]i[/b][/i][/font])"
        
    Case "(R¡Ï¦¨)"
        t = "([font=Times New Roman][i]r[/i]¡Ï[i]¦È[/i][/font])"
        
    '**************** ±¸ÓÃ
    'Case "[]"
    '    t = SetURLtoImg("")
        
    'Case "()"
    '    t = SetURLtoImg("")
        
        
    '****************
    Case Else
        t = ImageTxt
End Select
GetImageTxt = t
End Function

Public Function SetURLtoImg(URL As String) As String
    SetURLtoImg = "[img]" & URL & "[/img]"
End Function

Public Function SetSub(Text As String) As String
    SetSub = "[sub]" & URL & "[/sub]"
End Function

Public Function SetSup(Text As String) As String
    SetSup = "[sup]" & URL & "[/sup]"
End Function

Public Function SetFormat(Text As String, Font As String, It As Boolean, Bold As Boolean, udLine As Boolean) As String

End Function


