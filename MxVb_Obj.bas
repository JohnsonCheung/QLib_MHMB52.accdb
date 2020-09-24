Attribute VB_Name = "MxVb_Obj"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Obj."
Const DoczP$ = "Prpp."
Const DoczPn$ = "PrpnL."
Enum eThwOpt: eThwEr: eNoThw: End Enum
Function IsEqObj(A, B) As Boolean: IsEqObj = ObjPtr(A) = ObjPtr(B): End Function
Function IsEqVar(A, B) As Boolean: IsEqVar = VarPtr(A) = VarPtr(B): End Function

Function IntoOy(Into, Oy)
Erase Into
Dim O, I
For Each I In Itr(Oy)
    PushObj Into, I
Next
End Function

Function LngyOyPrp(Oy, Prpp$) As Long()
LngyOyPrp = CvLngy(IntoOyPrp(LngyEmp, Oy, Prpp))
End Function

Function IntoOyPrp(Into, Oy, Prpp$)
Dim O: O = AyNw(Into)
Dim Obj: For Each Obj In Itr(Oy)
    Push O, Opv(Obj, Prpp)
Next
IntoOyPrp = O
End Function

Sub ChkNothing(A, Fun$)
If IsNothing(A) Then Thw Fun, "Given object is nothing"
End Sub

Function Objn$(Obj)
If IsNothing(Obj) Then
    Objn = "Objn(Nothing)"
Else
    Objn = Obj.Name
End If
End Function
Function Objv(Obj)
On Error GoTo X
Objv = Obj.Value
Exit Function
X: Objv = "Er(" & Err.Description & ")"
End Function
