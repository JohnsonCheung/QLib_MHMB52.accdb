Attribute VB_Name = "MxIde_Md_TimMdy"
Option Compare Text
Option Explicit
Const TimMdy As Date = #9/18/2020 9:53:05 PM#


Sub EnsTimMdyMdn(Mdn$):          EnsTimMdyM Md(Mdn):                                    End Sub
Sub EnsTimMdyM(M As CodeModule): EnsCnstln M, FmtQQ("Const TimMdy As Date = #?#", Now): End Sub
Sub EnsTimMdyMC():               EnsTimMdyM CMd:                                        End Sub
Sub EnsTimMdy():                 EnsTimMdyM CMd:                                        End Sub
Function aTimMdyS(Dcl$()) As Date
Const Pfx$ = "Const TimMdy As Date = "
Dim A$: A = EleFstLik(Dcl, Pfx & "*"): If A = "" Then Exit Function
Dim B$: B = RmvFstLas(RmvPfx(A, Pfx))
aTimMdyS = B
End Function
Function aTimMdyMC() As Date:               aTimMdyMC = aTimMdyM(CMd):     End Function
Function aTimMdyM(M As CodeModule) As Date:  aTimMdyM = aTimMdyS(SrcM(M)): End Function
