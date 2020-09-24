Attribute VB_Name = "MxXls_XlsTy"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Ty."
Enum eXlsTy: eXlsTyNbr: eXlsTyTxt: eXlsTyTorN: eXlsTyDte: eXlsTyBool: End Enum 'Deriving(Enmm4 Enmt)
Public Const EnmttmlXlsTy$ = "N T TorN D B"
Public Const EnmqssXlsTy$ = "eXlsTy? Nbr Txt TorN Dte Bool"
Function eXlsTyyzCsl(CslXlsTy$) As eXlsTy()
Dim Ay$(): Ay = AmTrim(Split(CslXlsTy, ","))
Dim U%: U = UB(Ay)
Dim O() As eXlsTy: ReDim O(U)
Dim J%: For J = 0 To U
    O(J) = EnmvXlsTyTxt(Ay(J))
Next
eXlsTyyzCsl = O
End Function
Function EnmtXlsTy$(E As eXlsTy):         EnmtXlsTy = EleMsg(EnmtxtyXlsTy, E):  End Function
Function EnmsXlsTy$(E As eXlsTy):         EnmsXlsTy = EleMsg(EnmsyXlsTy, E):    End Function
Function EnmvXlsTy(S$) As eXlsTy:         EnmvXlsTy = IxEle(EnmsyXlsTy, S):     End Function
Function EnmvXlsTyTxt(Txt$) As eXlsTy: EnmvXlsTyTxt = IxEle(EnmtxtyXlsTy, Txt): End Function
Function EnmtxtyXlsTy() As String()
Static X$(): If Si(X) = 0 Then X = Tmy(EnmttmlXlsTy)
EnmtxtyXlsTy = X
End Function
Function EnmsyXlsTy() As String()
Static X$(): If Si(X) = 0 Then X = NyQss(EnmqssXlsTy)
EnmsyXlsTy = X
End Function
