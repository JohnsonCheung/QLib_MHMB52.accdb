Attribute VB_Name = "MxVb_Str_Seed"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Seed."
Enum ePjKd: ePjKdFba: ePjKdFxa: End Enum: Public Const EnmmPjKd$ = "ePjKd? Fba Fxa"
Sub ChkIsPthSrc(Pth$, Fun$)
If Not IsPthSrc(Pth) Then Thw Fun, "Given @Pth is not a Src path.  (A PthSrc should under a fdr of name [.src])", "Given-Pth", Pth
End Sub

Function DistPthiP$():                      DistPthiP = DistPthi(PthSrcPC):                          End Function
Function DistPthi$(PthSrc$):                 DistPthi = PthInst(DistPth(PthSrc)):                    End Function ':DistPthi: :Pthi ! #Dist-Pthi#
Function DistPth$(PthSrc$):                   DistPth = PthEns(Ffnn(Fdr(PthPar(PthSrc))) & ".dist"): End Function
Function DistFbaiP$():                      DistFbaiP = DistFbai(PthSrcPC):                          End Function
Function DistiFbaP$(P As VBProject):        DistiFbaP = DistFbai(PthSrc(Pjf(P))):                    End Function
Function DistPthP$():                        DistPthP = DistPth(PthSrcPC):                           End Function 'Distribution Path
Function DistFbai$(PthSrc$):                 DistFbai = DistPthi(PthSrc) & DistFn(PthSrc, ePjKdFba): End Function
Function DistFn$(PthSrc$, K As ePjKd):         DistFn = Ffnn(PjfPthSrc(PthSrc)) & ExtPjKd(K):        End Function
Function PjfPthSrc$(PthSrc$):               PjfPthSrc = RmvFst(PthUpN(PthSrc, 2)):                   End Function
Function DistFxaInst$(PthSrc$):           DistFxaInst = DistPthi(PthSrc) & DistFn(PthSrc, ePjKdFxa): End Function
Function DistFxaInstPC$():              DistFxaInstPC = DistFxaInstP(CPj):                           End Function
Function DistFxaInstP$(P As VBProject):  DistFxaInstP = DistFxaInst(PthSrcP(P)):                     End Function
Function ExtPjKd$(K As ePjKd)
Select Case True
Case K = ePjKdFba: ExtPjKd = ".accdb"
Case K = ePjKdFxa: ExtPjKd = ".xlsa"
Case Else: ThwEnm CSub, K, EnmmPjKd
End Select
End Function


Private Sub B_DistFxaInst()
Dim PthSrc1$
GoSub T0
Exit Sub
T0:
    PthSrc1 = PthSrcPC
    Ept = "C:\Users\user\Documents\Projects\Vba\QLib\.Dist\QLib(002).xlam"
    GoTo Tst
Tst:
    Act = DistFxaInst(PthSrc1)
    C
    Return
End Sub
