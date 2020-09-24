Attribute VB_Name = "MxVb_Obj_Opv"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Obj_Opv."

Function Opvy(Obj, Prppy$()) As Variant()
Const CSub$ = CMod & "Opvy"
If IsNothing(Obj) Then Inf CSub, "Given object is nothing", "Prppy", Prppy: Exit Function
Dim Prpp: For Each Prpp In Prppy
    Push Opvy, Oppv(Obj, Prpp)
Next
End Function

Function OpvDi(Obj As Object, Prppy$()) As Dictionary
'OpvDi:: :Dic #Object-Property-Value-Dic#
Set OpvDi = New Dictionary
Dim Prpp: For Each Prpp In Prppy
    OpvDi.Add Prpp, Opv(Obj, Prpp)
Next
End Function

Function Oppv(Obj, Prpp) ' return :Opv by @Prpp
'Opv:Cml :Var #Obj-Prp-Val#
'Prpp:: :S #Property-Path#
Dim Py$(): Py = SplitDot(Prpp)
Dim U%: U = UB(Py)
Dim O
    Set O = Obj
    Dim J%: For J = 0 To U - 1     ' U-1 is to skip the last Pth-Seg
        Set O = Opv(O, Py(J)) ' in the middle of each path-seg, they must be object, so use [Set O = ...] is OK
    Next
Asg Opv(O, Py(U)), _
    Oppv      ' Last Prp may be non-object, so must use 'Asg'
End Function

Function Opv(O, P): Asg CallByName(O, P, VbGet), Opv: End Function ' return :Opv
