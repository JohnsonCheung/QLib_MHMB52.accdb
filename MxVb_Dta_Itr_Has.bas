Attribute VB_Name = "MxVb_Dta_Itr_Has"
Option Compare Text
Const CMod$ = "MxVb_Dta_Itr_Has."
Option Explicit

Function HasItpv(Itr, P, V) As Boolean
Dim Obj: For Each Obj In Itr
    If Opv(Obj, P) = V Then HasItpv = True: Exit Function
Next
End Function
Function HasItppv(Itr, Prpp, V) As Boolean
Dim Obj: For Each Obj In Itr
    If Oppv(Obj, Prpp) = V Then HasItppv = True: Exit Function
Next
End Function

Function HasTruePrpp(Itr, Prpp) As Boolean
Dim I: For Each I In Itr
    If Opv(CvObj(I), Prpp) Then HasTruePrpp = True: Exit Function
Next
End Function

Function HasTruePrp(Itr, Prpp) As Boolean
Dim I: For Each I In Itr
    If Opv(I, Prpp) Then HasTruePrp = True: Exit Function
Next
End Function
