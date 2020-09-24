Attribute VB_Name = "MxDao_RsPrt"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_RsPrt."

Function CslRsFny$(R As Dao.Recordset, Fny$()): CslRsFny = Csl(DrRsFny(R, Fny)): End Function
Function CslSy$(Sy$()):                            CslSy = JnCma(AmQuoDbl(Sy)):  End Function
Function CslRs$(R As Dao.Recordset):               CslRs = Csl(AvItr(R.Fields)): End Function
Function CsyRs(R As Dao.Recordset) As String()
PushI CsyRs, CslSy(FnyRs(R))
R.MoveFirst
While Not R.EOF
    PushI CsyRs, CslRs(R)
    R.MoveNext
Wend
End Function

Function CsyRsFf(R As Dao.Recordset, FF$) As String(): CsyRsFf = CsyRsFny(R, FnyFF(FF)): End Function
Function CsyRsFny(R As Dao.Recordset, Fny$()) As String()
R.MoveFirst
While Not R.EOF
    PushI CsyRsFny, CslRsFny(R, Fny)
    R.MoveNext
Wend
End Function

Function LyJnRs(R As Dao.Recordset, Optional Sep$ = " ", Optional FF$) As String()
With R
    Push LyJnRs, Join(FnyRs(R), Sep)
    While Not .EOF
        PushI LyJnRs, LnRs(R, Sep)
        .MoveNext
    Wend
End With
End Function

Function FmtRsFf(R As Dao.Recordset, FF$) As String(): FmtRsFf = MsgyNNAv(FF, DrRsFf(R, FF)): End Function
Function FmtRs(R As Dao.Recordset) As String():          FmtRs = MsgyNyAv(FnyRs(R), DrRs(R)): End Function
