Attribute VB_Name = "MxDta_Da_Fmt_FmtJnDy"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Fmt_FmtJnDy."

Function LyDySpc(Dy()) As String(): LyDySpc = LyDy(Dy, " "): End Function
Function LyDyDot(Dy()) As String(): LyDyDot = LyDy(Dy, "."): End Function
Function LyDy(Dy(), Optional Sep$) As String()  'Ret: :Ly by joining each :Dr in @Dy by @Sep
Dim Dr: For Each Dr In Itr(Dy)
    PushI LyDy, RTrim(Jn(Dr, Sep))
Next
End Function
