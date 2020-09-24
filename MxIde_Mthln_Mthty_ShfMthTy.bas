Attribute VB_Name = "MxIde_Mthln_Mthty_ShfMthTy"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthln_Mtht_ShfMthTy."

Function ShfShtMdy$(OLn$):     ShfShtMdy = ShtMdy(ShfMdy(OLn)):    End Function
Function ShfShtMthTy$(OLn$): ShfShtMthTy = ShtMthTy(ShfMtht(OLn)): End Function
Function ShfMdy$(OLn$)
ShfMdy = Mdy(OLn)
OLn = LTrim(RmvPfx(OLn, ShfMdy))
End Function

Function ShfMthKd$(OLn$)
Dim T$: T = ShfMtht(OLn)
If T = "" Then Exit Function
ShfMthKd = Mthkd(T)
End Function

Function TakMthkd$(Ln): TakMthkd = PfxPfxySpc(Ln, MthKdy): End Function
Function TakMtht$(Ln):   TakMtht = PfxPfxySpc(Ln, MthTyy): End Function
Function RmvMtht$(Ln):   RmvMtht = RmvPfxySpc(Ln, MthTyy): End Function
