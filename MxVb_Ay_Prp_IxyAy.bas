Attribute VB_Name = "MxVb_Ay_Prp_IxyAy"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_ToIxy."

Function IxyPatn(Ay, Patn$) As Long(): IxyPatn = IxyRx(Ay, Rx(Patn)): End Function
Function IxyRx(Ay, R As RegExp) As Long() '
If Si(Ay) = 0 Then Exit Function
Dim I, J&
For Each I In Ay
    If R.Test(I) Then Push IxyRx, J
    J = J + 1
Next
End Function

Function InySubayThw(Ay, Subay) As Integer()
Dim F: For Each F In Subay
    PushI InySubayThw, IxMust(Ay, F)
Next
End Function

Function InySubssThw(Fny$(), Subff$) As Integer(): InySubssThw = InyFnySub(Fny, Tmy(Subff)): End Function
Function InyDrsCc(D As Drs, CC$) As Integer():        InyDrsCc = InySubssThw(D.Fny, CC):     End Function
Function InyFnySub(Fny$(), SubFny$()) As Integer()
Dim F: For Each F In Itr(SubFny)
    PushI InyFnySub, IxEle(Fny, F)
Next
End Function

Function InyCnoss(Cnoss$) As Integer(): InyCnoss = WCnossInto(Lngy, Cnoss): End Function
Function IxyCnoss(Cnoss$) As Long():    IxyCnoss = WCnossInto(Inty, Cnoss): End Function
Private Function WCnossInto(Into, Cnoss$)
Dim A$(): A = SySs(Cnoss)
Dim O: O = Into: Erase O
If Si(A) > 0 Then
    ReDim O(UB(A))
    Dim J&, Cno: For Each Cno In A
        O(J) = Cno - 1 ' Cno - 1 is to convert Cno to Ix
        J = J + 1
    Next
End If
WCnossInto = O
End Function

Function IxyEley(Ay, Eley) As Long() ' #Ixy-Eley# return ix of each ele of @Eley within @Ay.  Thw if @SubSy-Ele not found in @Ay
Dim S: For Each S In Itr(Eley)
    PushI IxyEley, IxEle(Ay, S)
Next
End Function
Function IyEley(Ay, Eley) As Integer() ' #Ixy-Eley# return ix of each ele of @Eley within @Ay.  Thw if @SubSy-Ele not found in @Ay
Dim S: For Each S In Itr(Eley)
    Dim I%: I = IxEle(Ay, S)
    If I >= 0 Then PushI IyEley, I
Next
End Function

Function IxyDup(Ay) As Long()
If IsEmpAy(Ay) Then Exit Function
Dim Dup()
Dim J&
For J = 0 To UB(Ay)
    If HasEle(Dup, Ay(J)) Then
        PushI IxyDup, J
    Else
        PushI Dup, J
    End If
Next
End Function
Function IxyEle(Ay, Ele) As Long()
Dim J&
Dim V: For Each V In Itr(Ay)
    If V = Ele Then PushI IxyEle, J
    J = J + 1
Next
End Function

Function IxyNw(U&) As Long()
If U >= 0 Then ReDim IxyNw(U)
End Function

Function IxyU(U&) As Long()
Dim O&(): O = IxyNw(U)
Dim J&: For J = 0 To U
    O(J) = J
Next
End Function
Function IxyEleDif(Ay1$(), Ay2$()) As Long()
If Si(Ay1) <> Si(Ay2) Then ThwPm CSub, "Si of Ay1 and Ay2 must be same", "Si(Ay1) Si(Ay2)", Si(Ay1), Si(Ay2)
Dim Ix&: For Ix = 0 To UB(Ay1)
    If Ay1(Ix) <> Ay2(Ix) Then
        PushI IxyEleDif, Ix
    End If
Next
End Function

Function IxiyPfxy(Sy$(), Pfxy$()) As Integer()
Dim J%: For J = 0 To UB(Sy)
    If HasPfxySpc(Sy(J), Pfxy) Then PushI IxiyPfxy, J
Next
End Function

