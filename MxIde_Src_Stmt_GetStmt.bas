Attribute VB_Name = "MxIde_Src_Stmt_GetStmt"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Stmt."

Private Sub B_StmtyyPC():                                 VcLyy StmtyyPC:     End Sub
Function StmtyyPC() As Variant():              StmtyyPC = StmtyyP(CPj):       End Function
Function StmtyyP(P As VBProject) As Variant():  StmtyyP = StmtyySrc(SrcP(P)): End Function

Private Sub B_StmtySrc()
GoSub Z2
Exit Sub
Z2:
    VcAy StmtySrc(SrcPC), "StmtySrc__Tst"
    Return
Z1:
    BrwAy StmtySrc(SrcMC), "StmtySrc__Tst"
    Return
End Sub
Function StmtySrc(Src$()) As String()
Dim C$(): C = Contlny(Src)
Dim L: For Each L In Itr(C)
    PushIAy StmtySrc, Stmty(L)
Next
End Function
Function StmtyySrc(Src$()) As Variant()
Dim C$(): C = Contlny(Src)
Dim L: For Each L In Itr(C)
    PushSomSi StmtyySrc, Stmty(L)
Next
End Function
Private Sub B_Stmty()
'GoSub T1
'GoSub T2
'GoSub T3
'GoSub T4
'GoSub T5
GoSub T6
'GoSub T7
'GoSub T8
'GoSub T9
'GoSub ZZ
Exit Sub
Dim Contln
ZZ:
    Vc Stmty(SrcPC)
    Return
T1:
    Contln = "Dim A$: B"
    Ept = Sy("Dim A$", "B")
    Pass "T1 Brk"
    GoTo Tst
T2:
    Contln = "Label: AAA"
    Ept = Sy("Label:", "AAA")
    Pass "T1 Label"
    GoTo Tst
T3:
    Ept = Sy()
    Contln = ""
    GoTo Tst
T4:
    Contln = "X "" Inp::(Inpn,Ffn)+"""
    Ept = Sy(Contln)
    GoTo Tst
T5:

    Contln = "PushI O, FmtQQ(""Dif at position: ?"", P)"
    Ept = Sy(Contln)
    GoTo Tst
T6:
    Dim A$: A = Replace("JrclnLno = FmtQQ('Jmp''?:?''", "'", vbQuoDbl)
    Contln = "Function JrclnLno$(Mdn$, Lno&, Ln): " & A
    Ept = Sy( _
        "Function JrclnLno$(Mdn$, Lno&, Ln)", _
        A)
    GoTo Tst
T7:
    Contln = Replace("JrclnLno = FmtQQ('Jmp''?:?''", "'", vbQuoDbl)
    Ept = Sy(Contln)
    GoTo Tst
T8:
    Contln = "Function A$():: End Function"
    Ept = Sy("Function A$()", "", "End function")
    GoTo Tst
T9:
    Contln = "Function IsParentLvl(Lvl) As Boolean: IsParentLvl = MMLvl > Lvl: End Function"
    Ept = Sy( _
        "Function IsParentLvl(Lvl) As Boolean", _
        "IsParentLvl = MMLvl > Lvl", _
        "End Function")
    GoTo Tst
Tst:
    Act = Stmty(Contln)
    C
    Return
End Sub

Function StmtyWhDim(Stmty$()) As String()
Dim Stmt: For Each Stmt In Itr(Stmty)
    If HasPfx(Stmt, "Dim ") Then
        PushI StmtyWhDim, Stmt
    End If
Next
End Function
Function StmtyPC() As String():       StmtyPC = StmtySrc(SrcPC): End Function
Function StmtyDisPC() As String(): StmtyDisPC = AwDis(StmtyPC):  End Function
Function Stmty(Contln) As String()
Dim L$
    L = RmvVmk(Contln): If L = "" Then Exit Function
Dim P%
    P = InStr(L, ":")
    If P = 0 Then
        PushNB Stmty, Trim(L)
        Exit Function
    End If
Dim O$()
    PushNB O, ShfVblbl(L, P)
Again:
    If L = "" Then GoTo X
    PushNB O, ShfStmt(L)
    GoTo Again
X:
    Stmty = O
Static Y As Boolean: If Not Y Then Y = True: Debug.Print "Stmty: Need to remove following Stmt...."
    Dim U%: U = UB(Stmty): If U = -1 Then GoTo M
    Dim J%: For J = 0 To UB(O)
        If Trim(O(J)) = "" Then Stop
        If ChrFst(O(J)) = " " Or ChrLas(O(J)) = " " Then Stop
    Next
    Exit Function
M:
ThwLgc CSub, "There is some non-Vmk in @Contln, but no Stmt is shifted out!", "@Contln", Contln
End Function

Private Sub B_StmtFst()
'GoSub T1
GoSub T2
Exit Sub
Dim Contln
T1:
    Contln = "Private Function W4SkuDesyZerOrNeg(Co As Byte, IntNm$) As String(): W4SkuDesyZerOrNeg = W4ErSkuDesyzSql(W4SqlZerOrNeg(Co, IntNm)): End Function"
    Ept = "Private Function W4SkuDesyZerOrNeg(Co As Byte, IntNm$) As String()"
    GoTo Tst
T2:
    Contln = "Function JnNB$(Ay, Optional Sep$ = """"): JnNB = Join(AwNB(Ay), Sep): End Function"
    Ept = "Function JnNB$(Ay, Optional Sep$ = """")"
    GoTo Tst
Tst:
    Act = StmtFst(Contln)
    Debug.Print Act
    Debug.Print Ept
    C
    Return
End Sub
Private Function ShfVblbl$(OLnTrimd$, PosColon%)
If PosColon = 0 Then Exit Function
If HasSpc(Left(OLnTrimd, PosColon - 1)) Then Exit Function
ShfVblbl = Left(OLnTrimd, PosColon)      'It if a vb-label
OLnTrimd = LTrim(Mid(OLnTrimd, PosColon + 1))
End Function
Function StmtFst$(Contln)
If Contln = "" Then Exit Function
Dim L$
    L = Trim(RmvVmk(Contln))
Dim P%: P = InStr(L, ":")
StmtFst = ShfVblbl(L, P): If StmtFst <> "" Then Exit Function
StmtFst = ShfStmt(L)
End Function

Private Function ShfStmt$(OLn$)
Dim P%: P = PosColonStmt(OLn): If P = 0 Then ShfStmt = Trim(OLn): OLn = "": Exit Function
ShfStmt = Trim(Left(OLn, P - 1))
OLn = LTrim(Mid(OLn, P + 1))
Static X As Boolean: If Not X Then X = True: Debug.Print "ShfStmt: needs to handle #11:11:11#"
'34
'56 PM#)
'If WWShfStmt = "56 PM#)" Or WWShfStmt = "34" Then Stop
End Function

Private Function PosColonStmt%(L$)
Dim B%: B = 1
Again:
    Dim P%: P = InStr(B, L, ":"): If P = 0 Then Exit Function
    If IsEven(NDblQuo(Left(L, P - 1))) Then
        PosColonStmt = P
        Exit Function
    End If
    B = P + 1
    GoTo Again
End Function
Private Sub B_PosColonStmt()
GoSub T1
GoSub T2
GoSub T3
Exit Sub
Dim L$, Pos%
T1:
    L = ": ASD"
    Ept = 1
    GoTo Tst
T2:
    L = "ASD: AA"
    Ept = 4         ' PosColonStmt does not handle of vb-label
    GoTo Tst
T3:
    '    1234567 89 0
    L = "ASD = "":"": AA" ' But handle [:] within Vstr
    Ept = 10
    GoTo Tst
Tst:
    Act = PosColonStmt(L)
    Debug.Assert Act = Ept
    Return
End Sub

