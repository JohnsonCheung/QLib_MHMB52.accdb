Attribute VB_Name = "MxApp_AppPth"
Option Compare Text
Option Explicit
Const CMod$ = "MxApp_AppPth."
Function PthInpApn$(Apn$)
PthInpApn = PthInp & Apn & "\"
Static P As Boolean: If Not P Then PthEns PthInpApn: P = True
End Function
Function PthOupApn$(Apn$)
PthOupApn = PthOup & Apn & "\"
Static P As Boolean: If Not P Then PthEns PthOupApn: P = True
End Function
Function PthWrkApn$(Apn$)
PthWrkApn = PthWrk & Apn & "\"
Static P As Boolean: If Not P Then PthEns PthWrkApn: P = True
End Function

Function PthInp$()
PthInp = "C:\Users\Public\Logistic\SAPData\"
Static P As Boolean: If Not P Then PthEns PthInp: P = True
End Function
Function PthWrk$()
PthWrk = PthInp & "Wrk\"
Static P As Boolean: If Not P Then PthEns PthWrk: P = True
End Function
Function PthOup$()
PthOup = "C:\Users\Public\Output\"
Static P As Boolean: If Not P Then PthEns PthOup: P = True
End Function
Function PthTp$():                 PthTp = PthTpP(CPj):                       End Function
Function PthTpP$(P As VBProject): PthTpP = PthEns(PthAssP(P) & "Templates\"): End Function

Sub BrwPthInp():        BrwPth PthInp:         End Sub
Sub BrwPthOup():        BrwPth PthOup:         End Sub
Sub BrwPthWrk():        BrwPth PthWrk:         End Sub
Sub BrwPthInpApn(Apn$): BrwPth PthInpApn(Apn): End Sub
Sub BrwPthOupApn(Apn$): BrwPth PthOupApn(Apn): End Sub

Sub BrwPthTp(): BrwPth PthTp: End Sub
