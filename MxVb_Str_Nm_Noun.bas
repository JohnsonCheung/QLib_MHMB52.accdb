Attribute VB_Name = "MxVb_Str_Nm_Noun"
Option Compare Text
Option Explicit

Function NounFunn$(Funn): NounFunn = Tm1(Mi4NavvM2("Fun " & Funn)): End Function
Function VerbSubn$(Subn): VerbSubn = CCmlFst(Subn):                 End Function
Function NounyFunny(Funny$()) As String()
Dim Funn: For Each Funn In Itr(Funny)
    PushI NounyFunny, NounFunn(Funn)
Next
End Function
