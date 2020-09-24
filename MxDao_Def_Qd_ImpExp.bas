Attribute VB_Name = "MxDao_Def_Qd_ImpExp"
Option Compare Text
Const CMod$ = "MxDao_Def_Qd_ImpExp."
Option Explicit
Sub ExpQryC(): ExpQry CDb: End Sub
Sub ExpQryPth(D As Database, PthTo$)
Dim Q As Dao.QueryDef: For Each Q In D.QueryDefs
    If Not HasPfx(Q.Name, "~sq_f") Then
        Debug.Print Now, "Exp Qry", Q.Name
        WrtStr Q.Sql, PthTo & Q.Name & ".sql"
    End If
Next
End Sub
Sub ExpQry(D As Database): ExpQryPth D, PthEns(PthSrc(D.Name)): End Sub

Sub ImpQry(D As Database, PthFm$)
Dim F: For Each F In Ffny(PthFm & F, "*.sql")
        Debug.Print Now, "Imp Qry", F
    D.QueryDefs.Append QdNw(Fnn(F), LinesFt(F))
Next
End Sub
