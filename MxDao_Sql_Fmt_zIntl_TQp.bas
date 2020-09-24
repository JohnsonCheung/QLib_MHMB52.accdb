Attribute VB_Name = "MxDao_Sql_Fmt_zIntl_TQp"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Sql_TQp."
Type TQp: Qpt As eQpt: Qpr As String: End Type ' Deriving(Ctor Ay)
Enum eQpt: eQptUpd: eQptFm: eQptInto: eQptSel: eQptSelDis: eQptSet: eQptWh: eQptOrd: eQptGp: eQptHav: eQptInrJn: eQptLeftJn: End Enum 'Deriving(Txt Str)
Public Const EnmttmlQpt$ = "Update From Into Select [Select Distinct] Set Where [Order By] [Group By] Having [Inner Join] [Left Join]"
Public Const EnmqssQpt$ = "eQpt? Upd Fm Into Sel SelDis Set Wh Ord Gp Hav InrJn LeftJn"

Sub PushTQp(A() As TQp, M As TQp): Dim N&: N = SiTQp(A): ReDim Preserve A(N): A(N) = M: End Sub
Function UbTQp&(A() As TQp): UbTQp = SiTQp(A) - 1: End Function
Function SiTQp&(A() As TQp): On Error Resume Next: SiTQp = UBound(A) + 1: End Function
Function EnmsyQpt() As String()
Static X$(): If Si(X) = 0 Then X = NyQss(EnmqssQpt)
EnmsyQpt = X
End Function
Function EnmtQpt$(E As eQpt): EnmtQpt = EleMsg(EnmtxtyQpt, E): End Function
Function EnmsQpt$(E As eQpt): EnmsQpt = EleMsg(EnmsyQpt, E):   End Function
Function EnmvQpt(S$) As eQpt: EnmvQpt = IxEle(EnmsyQpt, S):    End Function
Function EnmtxtyQpt() As String()
Static X$(): If Si(X) = 0 Then X = Tmy(EnmttmlQpt)
EnmtxtyQpt = X
End Function
Function RepTQp$(Q As TQp): RepTQp = EnmsQpt(Q.Qpt) & " " & Q.Qpr: End Function
Function StrTQp$(Q As TQp): StrTQp = EnmtQpt(Q.Qpt) & " " & Q.Qpr: End Function
