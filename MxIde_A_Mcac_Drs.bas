Attribute VB_Name = "MxIde_A_Mcac_Drs"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_A_Mcac_Drs."
Function DrsTMthCacPjn(D As Database, Pjn$, Mdny$()) As Drs
Dim O As Drs
Dim J%, N: For Each N In Itr(Mdny)
    Dim R As Dao.Recordset: Set R = Rs(D, FmtQQ("Select CmpTy,Mdl from Md where Pjn='?' and Mdn='?'", Pjn, N))
    Dim Src$(): Src = SplitCrLf(CStr(R!Mdl))
    Dim T$: T = R!CmpTy
    Stop
    Dim Dr(): 'Dr = DrMdn(Pjn, T, N)
    Stop 'O = DrsAdd(O, DrsTMthcS(Src, Dr))
    J = J + 1
Next
DrsTMthCacPjn = O
End Function

Function MdlTbMdP$(Mdn): MdlTbMdP = MdlTbMd(CurrentDb, CPjn, Mdn): End Function

Function MdlTbMd$(D As Database, Pjn$, Mdn)
Dim B$: B = FmtQQ("Pjn='?' and Mdn='?'", Pjn, Mdn)
MdlTbMd = ValTF(D, "Md.Mdl", B)
End Function
