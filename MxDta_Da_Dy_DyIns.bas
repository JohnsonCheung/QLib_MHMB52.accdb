Attribute VB_Name = "MxDta_Da_Dy_DyIns"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Dy_DyIns."

Function DyInsDc(Dy(), Dc, Optional CixBef% = 0) As Variant()
Const CSub$ = CMod & "DyInsDc"
If Si(Dy) <> Si(Dc) Then Thw CSub, "Si(Dy) should eq Si(Dc)", "Si(Dy) Si(Dc)", Si(Dy), Si(Dc)
Dim Dr, Ix&: For Each Dr In Itr(Dy)
    PushI DyInsDc, AyIns(Dr, Dc(Ix), CixBef)
    Ix = Ix + 1
Next
End Function
Function DyInsDcV2(Dy(), V1, V2, Optional CixBef%) As Variant():         DyInsDcV2 = DyInsDcDr(Dy, Array(V1, V2)):         End Function
Function DyInsDcV3(Dy(), V1, V2, V3, Optional CixBef%) As Variant():     DyInsDcV3 = DyInsDcDr(Dy, Array(V1, V2, V3)):     End Function
Function DyInsDcV4(Dy(), V1, V2, V3, V4, Optional CixBef%) As Variant(): DyInsDcV4 = DyInsDcDr(Dy, Array(V1, V2, V3, V4)): End Function
Function DyInsDcDr(Dy(), Dr, Optional CixBef%) As Variant()
Dim IDr: For Each IDr In Itr(Dy)
    PushI DyInsDcDr, AyInsAy(Dr, IDr, CixBef)
Next
End Function

Function DrsInsDc(A As Drs, Coln$, V) As Drs
DrsInsDc = Drs( _
    Fny:=SySyEle(A.Fny, Coln), _
    Dy:=DyAddDc(A.Dy, V) _
    )
End Function
Function DrsInsDc2V(D As Drs, P12$, V1, V2, Optional IsAtEnd As Boolean, Optional FldnBef$) As Drs: DrsInsDc2V = DwInsFf(D, P12, DyInsDcV2(D.Dy, V1, V2)):       End Function
Function DrsInsDc3V(D As Drs, P123$, V1, V2, V3) As Drs:                                            DrsInsDc3V = DrsFfAdd(D, P123, DyInsDcV3(D.Dy, V1, V2, V3)): End Function
