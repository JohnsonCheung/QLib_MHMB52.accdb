Attribute VB_Name = "MxIde_Md_Mdn_Gn"
Option Compare Text
Const CMod$ = "MxIde_Md_Mdn_Gn."
Option Explicit


Private Function WRxMdGn() As RegExp
Static X As RegExp
'If IsNothing(X) Then Set X = Rx("Mx[A-Z][a-z]+[A-Z][a-z0-9]+")
If IsNothing(X) Then Set X = Rx("^Mx[A-Z][a-z]+")
Set WRxMdGn = X
End Function

Function MdGnyPC() As String():              MdGnyPC = MdGnyP(CPj):             End Function
Function MdGnyP(P As VBProject) As String():  MdGnyP = AySrtQ(MdGny(MdnyP(P))): End Function
Function MdGny(Mdny$()) As String():           MdGny = AwRx(Mdny, WRxMdGn):     End Function
Function MdGn$(Mdn):                            MdGn = SsubRx(Mdn, WRxMdGn):    End Function
