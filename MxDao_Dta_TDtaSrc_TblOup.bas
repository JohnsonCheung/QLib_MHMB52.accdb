Attribute VB_Name = "MxDao_Dta_TDtaSrc_TblOup"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Dta_TDtaSrc_TblOup."

Private Sub B_TDtaSrcTblOup()
Dim A As TDtaSrc: A = TDtaSrcTblOup(MH.FcSamp.DbOup)
End Sub
Function TDtaSrcTblOup(D As Database) As TDtaSrc ' ret a :TDtaSrc from @D for all OTb (Tbn like @*)
With TDtaSrcTblOup
    .Fm = TDtaSrcFm(D.Name, "Tbl")
    Dim T: For Each T In Itr(TnyOup(D))
        PushTF .TF, TF(T, Fny(D, T))
    Next
End With
End Function
