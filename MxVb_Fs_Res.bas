Attribute VB_Name = "MxVb_Fs_Res"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_RES."

Sub EdtRES(SegFn$):                         VcFt resFfn(SegFn):          End Sub
Function resDrs(SegFn$) As Drs:    resDrs = DrsFt(resFfn(SegFn)):        End Function
Function resFfn$(SegFn$):          resFfn = resHom & SegFn:              End Function
Function resFfnEns$(SegFn$):    resFfnEns = FfnEnsPthAll(resFfn(SegFn)): End Function

Function resHom$()
Static H$: If H = "" Then H = PthAddFdrEns(PthAssPC, ".res")
resHom = H
End Function

Function Resl$(SegFn$):                                        Resl = LinesFt(resFfn(SegFn)):             End Function
Function resPth$(Pseg$):                                     resPth = PthEnsSfx(resHom & Pseg):           End Function
Function resPthzEns$(Pseg$):                             resPthzEns = PthEnsAll(resPth(Pseg)):            End Function
Function resS12y(SegFn$) As S12():                          resS12y = S12yS12lny(Resy(SegFn)):            End Function
Function Resy(SegFn$) As String():                             Resy = SplitCrLf(Resl(SegFn)):             End Function
Function sampFfn$(Fn$):                                     sampFfn = sampHom & Fn:                       End Function
Function sampFfny() As String():                           sampFfny = Ffny(sampHom):                      End Function
Function sampFnay() As String():                           sampFnay = Fnay(sampHom):                      End Function
Function sampHom$():                                        sampHom = resPthzEns("Sample"):               End Function
Sub WrtRes(Resy$(), SegFn$, Optional OvrWrt As Boolean):              WrtAy Resy, resFfn(SegFn), OvrWrt:  End Sub
Sub WrtRESDrs(D As Drs, SegFn$):                                      WrtDrs D, resFfn(SegFn):            End Sub
Sub WrtResl(Resl$, SegFn$, Optional OvrWrt As Boolean):               WrtStr Resl, resFfn(SegFn), OvrWrt: End Sub
