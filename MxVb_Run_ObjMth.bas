Attribute VB_Name = "MxVb_Run_ObjMth"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Run_ObjMth."
Sub RunItoMth(Ito, ObjMth)
Dim Obj As Object: For Each Obj In Ito
    CallByName Obj, ObjMth, VbMethod
Next
End Sub

Sub RunOyMth(Oy, ObjMth)
Dim Obj: For Each Obj In Itr(Oy)
    CallByName Obj, ObjMth, VbMethod
Next
End Sub
