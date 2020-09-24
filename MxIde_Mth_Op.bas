Attribute VB_Name = "MxIde_Mth_Op"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_Op."
Private Sub B_DltMth()
Const CSub$ = CMod & "B_DltMth"
GoSub T1
'GoSub ZZ
Exit Sub
Dim M As CodeModule, Mthn$
T1:
    DltMth CMd, "sdfdfdf"
    Return
ZZ:
Mthn = "YYRmv1"
Dim Bef$(), Aft$()
Crt:
        Set M = MdTmp
        AppCdl M, LinesVbl("|'sdklfsdf||'dsklfj|Property Get YYRmv1()||End Property||Function YYRmv2()|End Function||'|Sub SetYYRmv1(V)|End Property")
Tst:
        Bef = SrcM(M)
        DltMth M, Mthn
        Aft = SrcM(M)

Insp:   Insp CSub, "DltMth Test", "Bef DltMth Aft", Bef, Mthn, Aft
Rmv:    RmvMd M
        Return
End Sub

Sub MovMth(Mthn, ToMdn)
MovMthM CMd, Mthn, Md(ToMdn)
End Sub

Sub MovMthM(Md As CodeModule, Mthn, MdTo As CodeModule)
CpyMth Mthn, Md, MdTo
DltMth Md, Mthn
End Sub

Function CdlEmpFun$(FunNm)
CdlEmpFun = FmtQQ("Function ?()|End Function", FunNm)
End Function

Function CdEmpSub$(Subn)
CdEmpSub = FmtQQ("Sub ?()|End Sub", Subn)
End Function

Sub AddSub(Subn)
AppCdl CMd, CdEmpSub(Subn)
JmpMth Subn
End Sub

Sub AddFun(FunNm)
AppCdl CMd, CdlEmpFun(FunNm)
JmpMth FunNm
End Sub
