Attribute VB_Name = "MxVb_Str_Nm_NmNxt"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_NmNxt."

Function NumNxt&(Nm$, Optional NDig% = 3)
NumNxt = NumCur(Nm, NDig) + 1
End Function
Function NumCur&(Nm$, Optional NDig% = 3)
If HasNumNxt(Nm, NDig) Then
    NumCur = Right(Nm, NDig)
End If
End Function
Function HasNumNxt(Nm$, Optional NDig% = 3) As Boolean
If Len(Nm) < 2 + NDig Then Exit Function
Dim R$: R = Right(Nm, NDig + 1)
If ChrFst(R) <> "_" Then Exit Function
If Not IsNumeric(RmvFst(R)) Then Exit Function
HasNumNxt = True
End Function
Function NmRmvNum$(Nm$, Optional NDig% = 3)
If HasNumNxt(Nm, NDig) Then
    NmRmvNum = RmvLasN(Nm, NDig + 1)
Else
    NmRmvNum = Nm
End If
End Function
Function NmNxt$(Nm$, Optional NDig% = 3)
'   If XXX, return XXX_001   '<-- # of zero depends on NDig
'   If XXX_nn, return XXX_mm '<-- mm is nn+1, # of digit of nn and mm depends on NDig
Const CSub$ = CMod & "NmNxt"
If Not IsBet(NDig, 1, 7) Then ThwPm CSub, "Should between 1 and 7", "NDig", NDig
NmNxt = NmRmvNum(Nm) & "_" & Pad0(NumNxt(Nm, NDig), NDig)
End Function
