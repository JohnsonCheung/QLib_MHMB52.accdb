Attribute VB_Name = "MxDao_Dbt_AlterBoolFld"
Option Explicit
Option Compare Text
Const CMod$ = "MxDao_Dbt_AlterBoolFld."
Sub AltBoolFld(D As Database, T$, F$, Optional StrTrue$ = "Y", Optional FalseStr$ = "N")
RenFld D, T, F, F & "(Bool)"
Dim L%: L = Max(Len(StrTrue), Len(FalseStr))
Runq D, FmtQQ("Alter table [?] Add Column [?] Text(?)", T, F, L)
Runq D, FmtQQ("Update [?] set [?]=IIf(IsNull([?]),'',IIf([?(Bool)],'?','?'))", T, F, F, F, StrTrue, FalseStr)
Runq D, FmtQQ("Alter table [?] Drop Column [?(Bool)]", T, F)
End Sub
