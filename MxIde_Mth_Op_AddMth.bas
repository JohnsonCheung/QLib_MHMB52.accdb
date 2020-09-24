Attribute VB_Name = "MxIde_Mth_Op_AddMth"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_Op_Add."

Sub AddMthSub(M As CodeModule, Subn, Cdl$, Optional IsPrv As Boolean)
M.AddFromString CdlSub(Subn, Cdl, IsPrv)
End Sub
Sub AddMthFun(M As CodeModule, Funn, TyChr$, AsRet$, Cdl$, Optional IsPrv As Boolean)
M.AddFromString CdlFun(Funn, Cdl, TyChr, AsRet, IsPrv)
End Sub

Function CdlSub$(Subn, Cdl$, Optional IsPrv As Boolean)
CdlSub = LinesApLn( _
    FmtQQ("?Sub ?() As String()", MdyPrv(IsPrv), Subn), _
    Cdl, _
    "End Sub")
End Function
Function CdlFun$(Funn, TyChr$, AsRet$, Cdl$, Optional IsPrv As Boolean)
CdlFun = LinesApLn( _
    FmtQQ("?Function ??()?", MdyPrv(IsPrv), Funn, TyChr, AsRet), _
    Cdl, _
    "End Function")
End Function
