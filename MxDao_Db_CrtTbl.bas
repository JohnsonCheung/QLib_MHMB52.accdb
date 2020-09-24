Attribute VB_Name = "MxDao_Db_CrtTbl"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_CrtTbl."

Sub CrttEmpFm(D As Database, T, FmTbl$)
Runq D, SqlIntoSelStarWhFalse(T, FmTbl)
End Sub

Sub CrttDrs(D As Database, T, Drs As Drs)
CrttEmpDrs D, T, Drs
InsTblDy D, T, Drs.Dy
End Sub
Sub CrttEmpDrs(D As Database, T, Drs As Drs)
Crtt D, T, WSpec(Drs)
End Sub
Private Function WSpec$(D As Drs)
Dim J%
Dim Fny$(): Fny = D.Fny
Dim Dy(): Dy = D.Dy
Dim UFny%: UFny = UB(Fny)
Dim Specy$(): ReDim Specy(UFny)
    For J = 0 To UFny
        Specy(J) = WSpecDc(DcDy(Dy, J))
    Next

Dim O$()
    ReDim O(UFny)
    For J = 0 To UFny
        O(J) = Fny(J) & " " & Specy(J)
    Next
WSpec = JnCmaSpc(O)
End Function
Private Function WSpecDc$(Dc())
Stop
End Function

Sub CrttDup(D As Database, T, FmTbl, KK$)
Dim K$, Jn$, Tmp$, J%
Tmp = "##" & Tmpn
K = QpFf(KK)
Dim Into$
D.Execute FmtQQ("Select Distinct ?,Count(*) as Cnt into [?] from [?] group by ? having Count(*)>1", K, Tmp, FmTbl, K)
D.Execute FmtQQ("Select x.* into [?] from [?] x inner join [?] a on ?", Into, FmTbl, Tmp, Jn)
Drp D, Tmp
End Sub

Sub CrttTbjnfld(D As Database, T, KK$, Jnfld$, Optional Sep$ = " ")
Dim Tar$, LisFld$
    Tar = T & "_Jn_" & Jnfld
    LisFld = Jnfld & "_Jn"
Stop 'Runq D, SqlSel_Fny_Into_Fm(Ny(KK), Tar, T)
AddFld D, T, LisFld, dbMemo
Dim KKIdx&(), JnFldIx&
    KKIdx = IxyEley(Fny(D, T), KK)
    JnFldIx = IxF(D, T, Jnfld)
InsTblDy D, T, DyJnFldKK(DyT(D, T), KKIdx, JnFldIx)
End Sub
