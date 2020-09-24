Attribute VB_Name = "MxDao_Dbt_OpRen"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Dbt_OpRen."

Sub RenTC(T, ToNm$):               RenT CDb, T, ToNm:          End Sub
Sub RenT(D As Database, T, ToNm$): D.TableDefs(T).Name = ToNm: End Sub

Sub RenTblPfx(D As Database, PfxFm$, PfxTo$)
Dim T As TableDef: For Each T In D.TableDefs
    If HasPfx(T.Name, PfxFm) Then
        T.Name = RplPfx(T.Name, PfxFm, PfxTo)
    End If
Next
End Sub

Sub RenTTStrPfx(D As Database, TT$, Pfx$)
Dim T: For Each T In FnyFF(TT)
    RenTblStrPfx D, CStr(T), Pfx
Next
End Sub
