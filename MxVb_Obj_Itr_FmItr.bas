Attribute VB_Name = "MxVb_Obj_Itr_FmItr"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_ItrOpTfm."

Function AvItr(Itr) As Variant(): AvItr = IntoItr(AvEmp, Itr): End Function

Function SyItrPrp(Itr, Prpp) As String()
Dim I: For Each I In Itr
    PushI SyItrPrp, Opv(I, Prpp)
Next
End Function

Function SyItv(Itr) As String()
Dim I: For Each I In Itr
    PushI SyItv, I.Value
Next
End Function

Function SyItr(Itr) As String(): SyItr = IntoItr(SyEmp, Itr): End Function

Function IntoItr(Into, Itr)
Dim O: O = Into: Erase O
Dim V: For Each V In Itr
    Push O, V
Next
IntoItr = O
End Function

Function IntoItrMap(Into, Itr, Map$)
Dim O: O = Into: Erase Into
Dim X: For Each X In Itr
    Push O, Run(Map, X)
Next
IntoItrMap = O
End Function

Function SyItp(Itr, Prpp) As String():      SyItp = WIntoItp(SyEmp, Itr, Prpp):   End Function
Function IntyItp(Itr, Prpp) As Integer(): IntyItp = WIntoItp(IntyEmp, Itr, Prpp): End Function

Private Function WIntoItp(Into, Itr, Prpp)
WIntoItp = AyNw(Into)
Dim Obj: For Each Obj In Itr
    Push WIntoItp, Oppv(Obj, Prpp)
Next
End Function

Private Sub B_AvItp():                        Vc AvItp(CPj.VBComponents, "CodeModule.CountOfLines"): End Sub
Function AvItp(Itr, P$) As Variant(): AvItp = WIntoItp(AvEmp, Itr, P):                               End Function
