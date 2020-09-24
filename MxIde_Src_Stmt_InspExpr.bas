Attribute VB_Name = "MxIde_Src_Stmt_InspExpr"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Stmt_InspExpr."

Private Sub B_StmtInsp()
Dim O$()
Dim Ix%: Ix = 1
Dim C As VBComponent: For Each C In CPj.VBComponents
    Dim N$: N = C.Name
    Dim S$(): S = SrcCmp(C)
    Dim L: For Each L In Itr(MthlnySrc(S))
        PushI O, Ix & " " & N & " " & L
        PushI O, vbTab & " " & StmtInsp(L, N)
        Ix = Ix + 1
    Next
Next
VcAy O
End Sub
Function StmtInsp$(Mthln, Mdn$)
With MsigLn(Mthln)
    Dim IsRet As Boolean: IsRet = IsRet = IsShtMthTyRet(.ShtMdy)
    If SiTArg(.Arg) = 0 And Not IsRet Then Exit Function
    Dim Varnn$:  Varnn = WSI_Varnn(.Arg, IsRet)
    Dim EprLis$: EprLis = WSI_EprLis(.Arg, IsRet, .Vt)
    Dim Fun$: Fun = Mdn & "." & .Mthn
    Const C$ = "Insp ""?"", ""?"", ""?"", ?"
    StmtInsp = FmtQQ(C, Fun, "Inspect", Varnn, EprLis)
End With
End Function
Private Sub WSI__StmtInsp(): End Sub
Private Function WSI_Varnn$(A() As TArg, IsRet As Boolean)
Dim N$: If IsRet Then N = "Ret "
WSI_Varnn = StrTrue(IsRet, "Ret ") & JnSpc(Argny(A))
End Function
Private Function WSI_EprLis$(A() As TArg, IsRet As Boolean, Ret As TVt)
Dim O$(): If IsRet Then PushI O, WSI_EprRet(IsRet, Ret) '#Insp-Epr-0
Dim J%: For J = 0 To UbTArg(A)
    PushI O, WSI_EprTArg(A(J))
Next
WSI_EprLis = JnCmaSpc(O)
End Function
Private Function WSI_EprTArg$(A As TArg)
If A.Vt.Tyc <> "" Then WSI_EprTArg = A.Argn: Exit Function '#Is-Stringable-TVt# is the TVt can be expressed in an expression
WSI_EprTArg = WSI_EprFmt(A.Argn, A.Vt)
End Function
Private Function WSI_EprRet$(IsShtMthTyRet As Boolean, Ret As TVt)
If IsShtMthTyRet Then WSI_EprRet = WSI_EprFmt("Ret", Ret)
End Function
Private Function WSI_EprFmt$(Nm$, T As TVt)
Dim O$
If T.Tyc <> "" Then WSI_EprFmt = O: Exit Function
Select Case T.Tyn
Case "Drs":        O = WEprCallFmtFun("FmtDrs", Nm)
Case "S12":        O = WEprCallFmtFun("FmtS12" & IIf(T.IsAy, "y", ""), Nm)
Case "CodeModule": O = WEprCallFmtFun("Mdn", Nm)
Case "Dictionary": O = WEprCallFmtFun("FmtDi", Nm)
Case Else: O = """NoFmtr(" & T.Tyn & ")"""
End Select
WSI_EprFmt = O
End Function

Private Function WEprCallFmtFun$(Funn$, Varn$): WEprCallFmtFun = Funn & "(" & Varn & ")": End Function
