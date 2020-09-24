Attribute VB_Name = "MxAcs_Rescu_IntlRescu"
Option Compare Text
Option Explicit
Function FbyCorrudC() As String():   FbyCorrudC = FbyCorrud(PthPC):                End Function
Function FbyCorrud(Pth) As String():  FbyCorrud = Ffny(Pth, "*(Corrupted).accdb"): End Function
Function FbCorrudFstC$()
Dim A$(): A = FbyCorrudC
If Si(A) > 0 Then FbCorrudFstC = A(0)
End Function
Function FbRescudFstC$(): FbRescudFstC = FbRescud(FbCorrudFstC): End Function
Function ItrFbCorrudC(): Asg ItrFbCorrud(PthPC), ItrFbCorrudC: End Function
Function ItrFbCorrud(Pth): Asg FbyCorrud(Pth), ItrFbCorrud: End Function
Function FbRescud$(FbCorrud): FbRescud = FfnRplFnsfx(FbCorrud, "(Corrupted)", "(Rescued)"): End Function
