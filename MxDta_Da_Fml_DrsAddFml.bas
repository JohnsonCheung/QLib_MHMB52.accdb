Attribute VB_Name = "MxDta_Da_Fml_DrsAddFml"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Fml_FmCol."

Function DrsAddFml(A As Drs, NewFld$, FunNm$, PmAy$()) As Drs
Const CSub$ = CMod & "DrsAddFml"
Dim Dy(): Dy = A.Dy
If Si(Dy) = 0 Then DrsAddFml = A: Exit Function
Dim Dr, U&, Ixy&(), Av()
Ixy = IxyEley(A.Fny, PmAy)
U = UB(A.Fny)
For Each Dr In Dy
    If UB(Dr) <> U Then Thw CSub, "Dr-Si is diff", "Dr-Si U", UB(Dr), U
    Av = AwIxy(Dr, Ixy)
    Push Dr, RunAv(FunNm, Av)
Next
DrsAddFml = Drs(SySyEle(A.Fny, NewFld), Dy)
End Function

Function DrsAddFmlly(A As Drs, FmlSy$()) As Drs
Dim O As Drs: O = A
Dim NewFld$, FunNm$, PmAy$(), Fml$, I
For Each I In Itr(FmlSy)
    Fml = I
    NewFld = Bef(Fml, "=")
    FunNm = IsBet(Fml, "=", "(")
    PmAy = SplitCma(BetBkt(Fml))
    O = DrsAddFml(O, NewFld, FunNm, PmAy)
Next
End Function
