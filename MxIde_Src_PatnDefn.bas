Attribute VB_Name = "MxIde_Src_PatnDefn"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_PatnDefn."
Enum eColon: eColonYes: eColonNo: End Enum

Private Sub B_RfDefy():                   VcAy RfDefy(SrclPC, eColonYes): End Sub ':Johnson
Function HasDefr(S) As Boolean: HasDefr = HasRx(S, RxRfDef):              End Function
Function Defn$(S):                 Defn = SsubRx(S, RxRfDef):             End Function ' #Def-name# a name between 2 hashChr
Function RxRfDef() As RegExp '#Defintion-Reference# /:xxx / or /:xxx$/
Static X As RegExp: If IsNothing(X) Then Set X = Rx("/:([A-Za-z][\w\.-]*$)/gm")
'Static X As RegExp: If IsNothing(X) Then Set X = Rx(":([A-Za-z][\w\.-]*) |:([A-Za-z][\w\.-]*)$", IsGlobal:=True)
Set RxRfDef = X
End Function
Function RfDefy(S$, Optional C As eColon) As String()
RfDefy = SsubyRx(S, RxRfDef)
If C = eColonNo Then RfDefy = AmRmvFstChr(RfDefy)
End Function
