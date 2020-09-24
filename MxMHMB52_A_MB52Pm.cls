VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_MB52Pm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit
Const PmnPthi$ = "MB52_InpPth"
Const PmnIsCpy1$ = "MB52_IsCpyToPth1"
Const PmnIsCPy2$ = "MB52_IsCpyToPth2"
Const PmnPthCpy2$ = "MB52_CpyToPth2"
Const PmnPthCpy1$ = "MB52_CpyToPth1"

Private WPthi$, WPthCpy2$, WPthCpy1$, WIsCpy1 As Boolean, WIsCpy2 As Boolean
Property Get Pthi$():             Pthi = WPthi:       End Property
Property Get PthCpy1$():          PthCpy1 = WPthCpy1: End Property
Property Get PthCpy2$():          PthCpy2 = WPthCpy2: End Property
Property Get IsCpy1() As Boolean: IsCpy1 = WIsCpy1:   End Property
Property Get IsCpy2() As Boolean: IsCpy2 = WIsCpy2:   End Property

Private Sub Class_Initialize()
RfhPm
End Sub
Function PthCpy1SelSet$(): PthCpy1SelSet = PmvSetIfC(PthSel(PthCpy1, "Select first path that MB52 report will copy to"), PmnPthCpy1):  End Function
Function PthCpy2SelSet$(): PthCpy2SelSet = PmvSetIfC(PthSel(PthCpy2, "Select second path that MB52 report will copy to"), PmnPthCpy2): End Function
Sub RfhPm()
WPthi = PmvC(PmnPthi)
WIsCpy1 = PmvC(PmnIsCpy1)
WIsCpy2 = PmvC(PmnIsCPy2)
WPthCpy2 = PmvC(PmnPthCpy2)
WPthCpy1 = PmvC(PmnPthCpy1)
End Sub
