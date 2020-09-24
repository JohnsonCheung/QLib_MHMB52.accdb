Attribute VB_Name = "MxIde_Src_SrcItm"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_SrcItm."
Private Sub B_SrcItm()
MsgBox SrcItm("Private Sub SrcItm")
End Sub

Private Sub B_IsVbItm()
MsgBox IsVbItm("Sub")
End Sub

Function IsVbItm(Itm) As Boolean
':SrcItm: :S ! One of :VbItmy
':VbItmy: :Ny ! One of {Function Sub Type Enum Property Dim Const}
IsVbItm = HasEle(VbItmy, Itm)
End Function

Sub ChkIsVbItm(SrcItm, Fun$)
If Not IsVbItm(SrcItm) Then Thw Fun, "@SrcItm should be SrcItm", "@SrcItm Vdt-SrcItm", SrcItm, JnSpc(VbItmy)
End Sub

Function SrcItm$(Ln)
Dim O$: O = Tm1(RmvMdy(Ln))
If IsVbItm(O) Then SrcItm = O
End Function

Function VbItmy() As String()
Static X As Boolean, Y
If Not X Then
    X = True
    Y = SySs("Function Sub Type Enum Property Dim Const Option Implements")
End If
VbItmy = Y
End Function

Private Sub B_VbItmySrc()
Brw VbItmySrc(SrcPC)
End Sub

Function VbItmySrc(Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushNBNDup VbItmySrc, SrcItm(L)
Next
End Function
