Attribute VB_Name = "MxDao_Idx_Crt"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Idx_Crt."
Type TbFny: Tbn As String: Fny() As String: End Type ' Deriving(Ctor Ay)
Sub PushTFny(O() As TbFny, M As TbFny): Dim N&: N = SiTbFny(O): ReDim O(N): O(N) = M: End Sub
Function SiTbFny&(Ay() As TbFny): SiTbFny = UbTbFny(Ay): End Function
Function UbTbFny&(Ay() As TbFny): On Error Resume Next: UbTbFny = UBound(Ay): End Function
Function TbFnyzFf(T, FF$) As TbFny: TbFnyzFf = TbFny(T, FnyFF(FF)): End Function
Function TbFny(Tbn, Fny) As TbFny
With TbFny
    .Tbn = Tbn
    .Fny = Fny
End With
End Function
Sub CrtPk(D As Database, T):            D.Execute SqlCrtPk(T):         End Sub
Sub CrtSk(D As Database, T, Skff$):     D.Execute SqlCrtSkFf(T, Skff): End Sub
Sub CrtUKey(D As Database, T, K$, FF$): D.Execute SqlCrtUKy(T, K, FF): End Sub
Sub CrtKey(D As Database, T, K$, FF$)
D.Execute SqlCrtKey(T, K, FF)
End Sub

Function SqlCrtPk$(T)
SqlCrtPk = FmtQQ("Create Index PrimaryKey on [?] (?Id) with Primary", T, T)
End Function

Function SqlCrtSk$(A As TbFny): SqlCrtSk = FmtQQ("Create unique Index SecondaryKey on [?] (?)", A.Tbn, JnCma(Tmy(A.Fny))): End Function

Function SqyCrtSk(A() As TbFny) As String()
Dim J%: For J = 0 To UbTbFny(A)
    PushI SqyCrtSk, SqlCrtSk(A(J))
Next
End Function

Function SqlCrtSkFf$(T, Skff$): SqlCrtSkFf = SqlCrtSk(TbFnyzFf(T, Skff)): End Function

Function SqlCrtUKy$(T, K$, FF$)
Stop '
End Function

Function SqlCrtKey$(T, K$, FF$)
Stop '
End Function
Function SqyCrtPk(Tny$()) As String()
Dim T: For Each T In Itr(Tny)
    PushI SqyCrtPk, SqlCrtPk(T)
Next
End Function


