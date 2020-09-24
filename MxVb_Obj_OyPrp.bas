Attribute VB_Name = "MxVb_Obj_OyPrp"
':PP: :Prpp-PP$ #Spc-Separated-Prpp# ! Each ele is a Prpp
':Prpp: :Dotn   #Prp-Pth# ! Prp-Pth-of-an-Object
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Obj_PP."
Function OyAdd(Oy1, Oy2)
Dim O: O = Oy1
PushObjAy O, Oy2
OyAdd = O
End Function

Function FstOyEq(Oy, Prpp, V):         Set FstOyEq = ItoFstPrpEq(Itr(Oy), Prpp, V): End Function
Function AvOyP(Oy, Prpp) As Variant():       AvOyP = IntoOyP(AvEmp, Oy, Prpp):      End Function

Function IntoOyP(Into, Oy, Prpp)
Dim O: O = Into: Erase O
Dim Obj: For Each Obj In Itr(Oy)
    Push O, Opv(Obj, Prpp)
Next
IntoOyP = O
End Function

Function IntyOyP(Oy, Prpp) As Integer(): IntyOyP = IntoOyP(IntyEmp, Oy, Prpp): End Function
Function SyOyP(Oy, Prpp) As String():      SyOyP = IntoOyP(SyEmp, Oy, Prpp):   End Function
Function OyeNy(Oy, ExlNy$()) ' #Object-Array-where-excludate-name-array#
Dim O: O = Oy: Erase O
Dim I: For Each I In Itr(O)
    If Not HasEle(ExlNy, I.Name) Then PushObj O, I
Next
End Function
Function OyeNothing(Oy)
OyeNothing = AyNw(Oy)
Dim Obj As Object
For Each Obj In Oy
    If Not IsNothing(Obj) Then PushObj OyeNothing, Obj
Next
End Function

Function OywNmPfx(Oy, NmPfx$)
Dim Obj, O
O = Oy: Erase O
For Each Obj In Itr(Oy)
    If HasPfx(Obj.Name, NmPfx) Then PushObj O, Obj
Next
OywNmPfx = O
End Function

Function OywNm(Oy, B As TWhNm)
Dim Obj, O
O = Oy: Erase O
For Each Obj In Itr(Oy)
    If HitWhNm(Obj.Name, B) Then PushObj OywNm, Obj
Next
End Function

Function OywPredXPTrue(Oy, XP$, P$)
Dim O, Obj As Object
O = Oy
Erase O
For Each Obj In Itr(Oy)
    If Run(XP, Obj, P) Then
        PushObj O, Obj
    End If
Next
OywPredXPTrue = O
End Function

Function OyItr(Itr) As Variant()
Dim O
For Each O In Itr
    PushObj OyItr, O
Next
End Function
Function OywIn(Oy, Prpp, InAy)
Dim Obj As Object, O
If Si(Oy) = 0 Or Si(InAy) Then OywIn = Oy: Exit Function
O = Oy
Erase O
For Each Obj In Itr(Oy)
    If HasEle(InAy, Opv(Obj, Prpp)) Then PushObj O, Obj
Next
OywIn = O
End Function

Function LyObjPP(Obj As Object, PP$) As String()
Dim Prpp: For Each Prpp In SySs(PP)
    PushI LyObjPP, Prpp & " " & Opv(Obj, Prpp)
Next
End Function

Private Sub B_OyP_Ay()
Dim CdPanAy() As CodePane
Stop
'CdPanAy = Oy(CPj.MdAy).PrpVy("CodePane", CdPanAy)
Stop
End Sub
Private Sub B_LyObjPP()
Dim Obj As Object, PP$
GoSub T0
Exit Sub
T0:
    Set Obj = New Dao.Field
    PP = "Name Type Size"
    GoTo Tst
Tst:
    Act = LyObjPP(Obj, PP)
    C
    Return
End Sub
