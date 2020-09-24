Attribute VB_Name = "MxVb_Dta_Itr_ItrWh"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_ItrWh."

Function IntoItwEq(Into, Itr, Prpp, V)
IntoItwEq = AyNw(Into)
Dim Obj: For Each Obj In Itr
    If Opv(Obj, Prpp) = V Then PushObj IntoItwEq, Obj
Next
End Function

Private Sub B_IntoItwLik()
Dim Into() As CodeModule
Dim Act() As CodeModule: Act = IntoItwLik(Into, CPj.VBComponents, "Name", "MxVbStr*")
Stop
End Sub

Function IntoItwLik(Into, Itr, Prpp, Lik$)
IntoItwLik = AyNw(Into)
Dim Obj: For Each Obj In Itr
    If Oppv(Obj, Prpp) Like Lik Then
        PushObj IntoItwLik, Obj
    End If
Next
End Function

Function IntoItwTPrp(OInto, Itr, TPrpp) As Variant()
IntoItwTPrp = AyNw(OInto)
Dim Obj: For Each Obj In Itr
    If Oppv(Obj, TPrpp) Then
        PushObj IntoItwTPrp, Obj
    End If
Next
End Function
