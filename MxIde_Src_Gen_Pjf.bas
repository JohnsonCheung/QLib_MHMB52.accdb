Attribute VB_Name = "MxIde_Src_Gen_Pjf"
':PthSrc:    :Pth ! #Src-Path#           is a :Pth.  Its fdr is a `{PjFn}.src`
':Distp:   :Pth ! #Distribution-Path#  is a :Pth.  It comes from :PthSrc by replacing .src to .dist
':Pthi: :Pth ! #Instance-Path#      of a @pth is any :TimNm :Fdr under @pth
':TimNm:   :Nm
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Gen_Pjf."

Sub GenFxaFmAcs(Fxa$, A As Access.Application)
Stop 'ExpSrcAcs A
Stop 'GenFxa Fxa, PthSrcA(A)
End Sub

Sub GenTmpFxaPC(): GenTmpFxaP CPj: End Sub
Sub GenTmpFxaP(P As VBProject)
Dim Pth$: Pth = PthSrcP(P)
Dim T$: Stop 'T = TqmpFxa
Stop 'GenFxaPC T, P
Stop 'GenTmpFxaFmPth = T
End Sub
Sub GenFxaPC(Fxa$, PthSrc$)
Const CSub$ = CMod & "GenFxaPC"
:                                   CrtFxa Fxa          ' <== Crt
Dim X As Excel.Application: Set X = XlsNw
Dim P As VBProject:     Stop '    Set P = XlsOpnXla(X, Fxa)
:                      Stop '             AddRfzPthSrc P, PthSrc ' <== Add Rf
:                                   Stop 'LoadBas P, PthSrc    ' <== Load Bas
X.Quit
Inf CSub, "Fxa is created", "Fxa", Fxa
End Sub

Sub GenFba(PthSrc$)
Const CSub$ = CMod & "GenFba"
Dim OPj As VBProject
Dim OFba$:     OFba = DistFbai(PthSrc)     '#Oup-Fba#
:                     DltFfnIf OFba
:                     CrtFb OFba                    ' <== Crt OFba


:                     Acs.OpenCurrentDatabase OFba
            Set OPj = PjAcs(Acs)
:                     ImpRfP OPj, PthSrc ' <== Add Rf
:                     ImpSrcP OPj, PthSrc             ' <== Load Bas
Dim F: For Each F In Itr(FfnyFrmPth(PthSrc))
    Dim N$: N = Ffnn(Ffnn(F))
:               Acs.LoadFromText acForm, N, F       ' <== Load Frm
Next
#If False Then
'Following code is not able to save
Dim CVbe As VBE: Set CVbe = Acs.VBE
Dim C As VBComponent: For Each C In Acs.VBE.ActiveVBProject.VBComponents
    C.Activate
    BoSavzV(CVbe).Execute
    Acs.Eval "DoEvents"
Next
#End If
MsgBox "Go access to save....."
Inf CSub, "Fba is created", "Fba", OFba
End Sub


Sub GenFbaFmCAcs()
GenFbaFmAcs Acs
End Sub

Sub GenFbaFmAcs(A As Access.Application)
Stop 'ExpSrcAcs A                       ' <== Exp
GenFba PthSrcP(PjAcs(A))
End Sub
