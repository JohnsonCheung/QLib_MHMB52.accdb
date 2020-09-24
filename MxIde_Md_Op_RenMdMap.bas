Attribute VB_Name = "MxIde_Md_Op_RenMdMap"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Md_Op_RenMdMap."
Enum eRfh: eRfhNo: eRfhYes: End Enum
Sub EdtMdnMap()
Dim Fx$: Fx = WFxRenMd
DltFfnIf Fx
Dim B As Workbook: Set B = WbEns(Fx)
W2EnsLo WsFst(B)
Dim L As ListObject: Set L = LoFst(WsFst(B))
SavWb B
Maxv B.Application
ActAppFx B.Application
End Sub
Private Sub W2EnsLo(S As Worksheet)
Const CSub$ = CMod & "W2EnsLo"
If S.ListObjects.Count > 0 Then ThwImposs CSub, "Worksheet should always a new workseet"
PutAyHori SySs("MdnOld MdnNew"), A1Ws(S)
Dim Mdny1$(): Mdny1 = MdnyPC
Dim Mdny2$(): Mdny2 = SySrtQ(Mdny1)
PutAyVert Mdny2, S.Range("A2")
LoRg S.Range("A1:B" & Si(Mdny2) + 1)
End Sub

Sub RenMdMap()
Dim P As VBProject: Set P = CPj
Dim S() As S12: S = WS12yRen
Dim J%: For J = 0 To UbS12(S)
    With S(J)
        P.VBComponents(.S1).Name = .S2
    End With
Next
End Sub
Private Sub B_WS12yRen(): BrwS12y WS12yRen: End Sub
Private Function WS12yRen() As S12()
Dim B As Workbook: Set B = WbFx(WFxRenMd)
Dim L As ListObject: Set L = LoFst(WsFst(B))
Dim O() As S12: O = S12yLo(L, 1, 2)
O = S12yRmvBlnkS2(O)
If SiS12(O) = 0 Then Exit Function
WChk O
WS12yRen = O
End Function
Private Sub WChk(OldNew() As S12)
Const CSub$ = CMod & "WChk"
Dim I$(), Oldy$(), Newy$()
Oldy = S1y(OldNew)
Newy = S2y(OldNew)
I = AyIntersect(Oldy, Newy): If Si(I) > 0 Then Thw CSub, "There are common mdny in @Oldy and @Newy", "[Common Mdny] [@Oldy (Old Mdny)] [@Newy (New Mdny)]", I, Oldy, Newy
Dim Mdny$(): Mdny = MdnyPC
I = AyIntersect(Mdny, Newy): If Si(I) > 0 Then Thw CSub, "There are common mdny in current mdny and @Newy", "[Common Mdny] [Current Mdny] [@Newy (New Mdny)]", I, Mdny, Newy
I = SyMinus(Oldy, Mdny): If Si(I) > 0 Then Thw CSub, "There are mdny in @Oldy which is not exist", "[Oldy not in existing Mdny] [Existing Mdny] [@Oldy (Old Mdny)]", I, Mdny, Oldy
I = WLen64(I): If Si(I) > 0 Then Thw CSub, "There new mdny with len>64", "[New Mdn > Len64]", I
End Sub
Private Function WLen64(Mdny$()) As String()
Dim Mdn: For Each Mdn In Itr(Mdny)
    If Len(Mdn) > 64 Then PushI WLen64, Mdn
Next
End Function


Private Function X_Oldy(L As ListObject) As String():   X_Oldy = DcStrLc(Lc(L, 1)):        End Function
Private Function WFxRenMd$():                         WFxRenMd = PthTmp & "RenMdMap.xlsx": End Function
