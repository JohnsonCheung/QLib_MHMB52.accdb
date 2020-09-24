Attribute VB_Name = "MxIde_Lis_LisJrc"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Lis_LisJrc."
Enum eUL: eULYes: eULNo: End Enum
Function sampFhtml$()
Dim Fhtml$: Fhtml = PthTmpFdr("SampHtml") & "A.html"
Const Cxt$ = "<html><p>ABC</p></html>"
If Not HasFfn(Fhtml) Then EnsFt Fhtml, Cxt
sampFhtml = Fhtml
End Function
Sub VwSMPHtml(): VwHtml sampFhtml: End Sub
Sub VwHtml(Fhtml$)
Shell FmtQQ("Cmd /C ""?""", Fhtml), vbHide
End Sub
Sub LisPj()
Dim A$()
    A = PjnyV(CVbe)
    D AmAddPfx(A, "ShwPj """)
D A
End Sub

Sub LisJrc(Patn$, Optional PatnssAndMd$, Optional UL As eUL): DmpAy JrcyPatnPC(Patn, PatnssAndMd, UL):     End Sub
Sub LisStop(Optional PatnssAndMd$, Optional UL As eUL):       DmpAy JrcyPatnPC("Stop '", PatnssAndMd, UL): End Sub

Private Sub B_Beta()
'ß()
End Sub
Sub VcJrc(Patn$, Optional PatnssAndMd$, Optional UL As eUL):  VcAy JrcyPatnPC(Patn, PatnssAndMd, UL), "VcJrc_":       End Sub
Sub BrwJrc(Patn$, Optional PatnssAndMd$, Optional UL As eUL): BrwAy JrcyPatnPC(Patn, PatnssAndMd, UL), "JrcyPatnPC_": End Sub
