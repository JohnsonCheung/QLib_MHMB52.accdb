Attribute VB_Name = "MxVb_Dta_Lpm_LpmFun"
Option Compare Database
Option Explicit
Private R As RegExp
Type TLpmBrk: DiSw As Dictionary: DiLpmnv As Dictionary: End Type
Function SwvLpm(P As TLpmBrk, Lpmn$) As Boolean: SwvLpm = P.DiSw.Exists(Lpmn): End Function
Function TLpmBrk(Lpm$) As TLpmBrk
Dim S() As S12: S = S12yLpm(Lpm)
Dim O As TLpmBrk
Set O.DiSw = DiNwSen
Set O.DiLpmnv = DiNwSen
Dim J%: For J = 0 To UbS12(S)
    With S(J)
        If .S2 = "" Then
            O.DiSw(.S1) = Empty
        Else
            O.DiLpmnv(.S1) = .S2
        End If
    End With
Next
TLpmBrk = O
End Function
Function Lpmv$(P As TLpmBrk, Lpmn$): Lpmv = StrDivIf(P.DiLpmnv, Lpmn): End Function
Private Sub B_S12yLpm()
GoSub ZZ
Exit Sub
Dim Lpm$
Dim S12y() As S12
ZZ:
    Lpm = "-Prv -Pub -AA BB CC -E"
    S12y = S12yLpm(Lpm)
    Debug.Print Lpm
    Dim J%: For J = 0 To UbS12(S12y)
        Debug.Print J; "["; S12y(J).S1; "] ["; S12y(J).S2; "]"
    Next
    Return
End Sub
Function RxayLpm(P As TLpmBrk, Lpmn$) As RegExp(): RxayLpm = RxAyPatnss(Lpmv(P, Lpmn)): End Function
Private Function S12yLpm(Lpm$) As S12()
If Lpm = "" Then Exit Function
If ChrFst(Lpm) <> "-" Then Thw CSub, "ChrFst of @Lpm parameter line must be [-]", "Lpm", Lpm
If IsNothing(R) Then Set R = Rx("/-\w+ /g")
Dim M As MatchCollection: Set M = R.Execute(Lpm & " ") 'Add Spc at end is due to just in case the last pm is a switch
'----
Dim Sy$(): Sy = SsubyMchColl(M)
Dim Iy%(): Iy = IyMchColl(M)
ChkIsEqSi Sy, Iy, "S12yLpm"
Dim U&: U = UB(Sy)
Dim J%: For J = 0 To U
    Dim S1$: S1 = RmvFstLas(Sy(J))
    Dim S2$
        Dim P%: P = Iy(J) + Len(Sy(J))
        If J = U Then
            S2 = Mid(Lpm, P)       '<==
        Else
            Dim L%
                L = Iy(J + 1) - P + 1
            S2 = Mid(Lpm, P, L)        '<==
        End If
    PushS12 S12yLpm, S12(S1, Trim(S2))
Next
End Function
