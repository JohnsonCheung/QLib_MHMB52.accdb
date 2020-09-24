Attribute VB_Name = "MxXls_Fea_Lof_Lofl"
':Lofl: :Lines #Lo-Fmtr-Lines#
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Fea_Lof_Lofl."
Function LoflLo$(L As ListObject):      LoflLo = L.Comment:                 End Function
Function LoflFbt$(Fb, T):              LoflFbt = LoflT(Db(Fb), T):          End Function
Sub SetLoflFbt(Fb, T, Lofl$):                    SetLoflT Db(Fb), T, Lofl:  End Sub
Sub SetLoflT(D As Database, T, Lofl$):           SetPvT D, T, "Lofl", Lofl: End Sub

Property Get LoflTC$(T): Stop '     LoflTC = Tbprpv(CDb, T): End Property
End Property

Property Get LoflT$(D As Database, T):       LoflT = PvT(D, T, "Lofl"):         End Function
Property Get LoflT1XX$(D As Database, T)

End Property
