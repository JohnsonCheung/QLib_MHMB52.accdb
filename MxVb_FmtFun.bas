Attribute VB_Name = "MxVb_FmtFun"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_FmtFun."
Function SeplnWdty$(W%(), F As eTblFmt)
Dim Dr()
    Dim I: For Each I In Itr(W)
        Push Dr, StrDup("-", I)
    Next
SeplnWdty = LnDr(Dr, QmkSepln(F))
End Function
