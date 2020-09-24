Attribute VB_Name = "MxVb_Dta_Itr_Fst"
Option Compare Text
Const CMod$ = "MxVb_Dta_Itr_Fst."
Option Explicit

Function ItvFst(Itr)
Dim Itv: For Each Itv In Itr
    Asg Itv, ItvFst
    Exit Function
Next
End Function

Function ItoFstPrpEq(Itr, Prpp, V) '#(ItvFst)-of-@Itr-with-@Prpp-value-eq-to-@V#
Dim Ito: For Each Ito In Itr
    Dim Pv: Pv = Oppv(Ito, Prpp)
    If Pv = V Then Set ItoFstPrpEq = Ito: Exit Function
Next
Set ItoFstPrpEq = Nothing
End Function

Function ItoFstPrpTrue(Itr, Prpp): ItoFstPrpTrue = ItoFstPrpEq(Itr, Prpp, True): End Function '#(ItvFst)-of-@Itr-with-@Prpp-value-eq-(True)#
Function ItoFstNm(Itr, Nm$):        Set ItoFstNm = ItoFstPrpEq(Itr, "Name", Nm): End Function '#(ItvFst)-of-@Itr-with-Nm-prp-eq-@Nm#
