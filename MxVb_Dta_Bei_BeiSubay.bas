Attribute VB_Name = "MxVb_Dta_Bei_BeiSubay"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Bei_BeiSubay."
Function BeiSubay(Ay, Subay) As Bei ' retun the Bei of first Blk of ele from @Ay.  Each ele of the Blk is having one of the value in @Subay
Dim B&: B = WBixSubay(Ay, Subay)
Dim E&: E = WEixSubay(Ay, B, Subay)
BeiSubay = Bei(B, E)
End Function
Private Function WBixSubay&(Ay, Subay)
Dim J&: For J = 0 To UB(Ay)
    If HasEle(Subay, Ay(J)) Then WBixSubay = J: Exit Function
Next
WBixSubay = -1
End Function
Private Function WEixSubay&(Ay, Bix&, Subay)
If Bix = -1 Then GoTo Ext
Dim J&: For J = Bix + 1 To UB(Ay)
    If Not HasEle(Subay, Ay(J)) Then WEixSubay = J - 1: Exit Function
Next
WEixSubay = UB(Ay)
Exit Function
Ext: WEixSubay = -1
End Function
