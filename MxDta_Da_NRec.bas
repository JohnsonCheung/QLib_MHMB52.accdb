Attribute VB_Name = "MxDta_Da_NRec"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_NRec."
Function NRecDyCne&(Dy(), Ci%, Cnev)
Dim Dr: For Each Dr In Itr(Dy)
    If Dr(Ci) <> Cnev Then NRecDyCne = NRecDyCne + 1
Next
End Function

Function NRecDyCeq&(Dy(), Ci%, Ceqv)
Dim Dr: For Each Dr In Itr(Dy)
    If Dr(Ci) = Ceqv Then NRecDyCeq = NRecDyCeq + 1
Next
End Function

Function NRecDrsCeq&(D As Drs, C$, Ceqv): NRecDrsCeq = NRecDyCeq(D.Dy, IxiEle(D.Fny, C), Ceqv): End Function
Function NRecDrsCne&(D As Drs, C$, Cnev): NRecDrsCne = NRecDyCne(D.Dy, IxiEle(D.Fny, C), Cnev): End Function
Function NRecDrs&(D As Drs):                 NRecDrs = Si(D.Dy):                                End Function
Function NRecDt&(D As Dt):                    NRecDt = Si(D.Dy) = 0:                            End Function
