Attribute VB_Name = "MxDao_Db_EndDte"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_EndDte."

Sub UpdEndDte(D As Database, T, EndDteFld$, BegDteFld$, GpFf)
Dim LasBegDte As Date
LasBegDte = DateSerial(2099, 12, 31)
Dim Q$
''Q = SqlSelFf_Ordff(Sy(BegDteFld, EndDteFld), T, BegDteFld)
Stop
With Rs(D, Q)
    While Not .EOF
        .Edit
        .Fields(EndDteFld).Value = LasBegDte
        .Update
        .MoveNext
    Wend
    .Close
End With
End Sub
