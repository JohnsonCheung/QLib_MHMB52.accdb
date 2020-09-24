Attribute VB_Name = "MxVb_Dta_Itr_ItrPrp"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Itr_Itn_S."

Function ItnFst$(Itr)
Dim I: For Each I In Itr
    ItnFst = Objn(I)
Next
End Function
Function IxItn&(Itr, Nm)
Dim O&, I: For Each I In Itr
    If I.Name = Nm Then IxItn = O
    O = O + 1
Next
End Function
Function Itn(Itr) As String(): Dim I: For Each I In Itr: PushI Itn, Objn(I): Next: End Function
Function HasItn(Itr, Nm) As Boolean
Dim Obj: For Each Obj In Itr
    If Opv(Obj, "Name") = Nm Then HasItn = True: Exit Function
Next
End Function
Function ItnWhPrpNB(Itr, Prpp) As String()
Dim I: For Each I In Itr
    If Opv(I, Prpp) <> "" Then PushI ItnWhPrpNB, Objn(I)
Next
End Function
Function ItnWhPrpTrue(Itr, Prpp) As String()
Dim I: For Each I In Itr
    If Oppv(I, Prpp) Then PushI ItnWhPrpTrue, Objn(I)
Next
End Function
Function ItnWhPrpFalse(Itr, Prpp) As String()
Dim I: For Each I In Itr
    If Not Oppv(I, Prpp) Then PushI ItnWhPrpFalse, Objn(I)
Next
End Function

Function ItnWhPrpBlnk(Itr, Prpp) As String()
Dim I: For Each I In Itr
    If Oppv(I, Prpp) = "" Then PushI ItnWhPrpBlnk, Objn(I)
Next
End Function

Function ItnTyn(Itr, Tyn$) As String()
Dim O: For Each O In Itr
    If TypeName(O) = Tyn Then
        PushI ItnTyn, Objn(O)
    End If
Next
ItnTyn = SySrtQ(ItnTyn)
End Function
