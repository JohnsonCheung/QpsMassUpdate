Attribute VB_Name = "DtaChg"
Option Explicit
Enum eFld     ' At Ro5
    eSupplier = 1
    eBrand = 2
    eProjNo = 3
    eQuoteDate = 4
    ePotentialQty = 5
    eRateUSD = 6
    eRateCHF = 7
    eRateJPY = 8
    eSku = 9
    eSkuCost = 10
    eAssWatchUSD = 11
    eAssWatchHKD = 12
    eCompleteWatchUSD = 13
    eCompleteWatchHKD = 14
    eSalesmanUSD = 15
    eSalesmanHKD = 16
    eOneTimeCost01 = 17
    eOneTimeCost01Rmk = 18
    eOneTimeCost02 = 19
    eOneTimeCost02Rmk = 20
    e_EleGpNNTot = 21
    e_EleGpNNEleNNRmk = 22
    e_EleGpNNEleNN = 23
    e_ChrGpNNEleNNChrNN = 24
End Enum
'=========================================
Enum eFldTy
    ePjQ = 1
    eOne = 2
    eSku = 3
    eCstVal = 4
    eCstRmk = 5
    eChr = 6
End Enum


Type TDtaChg
    FldTy As eFldTy     'One Sku CstVal CstRmk Chr
    FldNm As String
    Cno As Integer
    Key As KeyDta       ' The Row# is stored in "Key"
    CostGp As String
    CostEle As String
    CharName As String
    CharCode As String  ' Use in update table-SkuCostChr
    OrgVal As Variant
    WrkVal As Variant
End Type

Type PjKey
    Pj As String
    QDte As Date
End Type

Private Enum eOptional
    eOpt = 1
    eMust = 2
End Enum
Dim O() As TDtaChg
Dim WrkCno As TCno
Dim OrgCno As TCno
Dim WrkSqv, OrgSqv
Dim WrkKey() As KeyDta

Function DtaChg_KeyDta(Ay() As TDtaChg) As KeyDta()
'From A_DcPush, return a unique KeyDta[]
Dim N&
    N = DtaChg_Sz(Ay)
If N = 0 Then Exit Function
Dim O() As KeyDta
    Dim J&
    ReDim O(N - 1)
    For J = 0 To N - 1
        O(J) = Ay(J).Key
    Next
DtaChg_KeyDta = O
End Function

Property Get DtaChg_Sz&(Ay() As TDtaChg)
On Error Resume Next
DtaChg_Sz = UBound(Ay) + 1
End Property

Property Get DtaChg_YellowAdr(Ay() As TDtaChg) As YellowAdr
Dim N&, O As YellowAdr, J&
    N = DtaChg_Sz(Ay)
If N = 0 Then Exit Property
ReDim O.C(N - 1)
ReDim O.R(N - 1)
For J = 0 To N - 1
    O.C(J) = Ay(J).Cno
    O.R(J) = Ay(J).Key.Rno + 6
Next
DtaChg_YellowAdr = O
End Property

Function Enm_FldTy(S$) As eFldTy
Dim O As eFldTy
Select Case S
Case "ePjQ":    O = eFldTy.ePjQ
Case "eOne":    O = eFldTy.eOne
Case "eSku":    O = eFldTy.eSku
Case "eCstVal": O = eFldTy.eCstVal
Case "eCstRmk": O = eFldTy.eCstRmk
Case "eChr":    O = eFldTy.eChr
Case Else: Er "Given {S} is a not in valid Enm-eFldTy-{MbrNmList}", S, "[ePjQ eOne eSku eCstVal eCstRmk eChr]"
End Select
Enm_FldTy = O
End Function

Function Enm_FldTy_ToStr(P As eFldTy)
Dim O$
Select Case P
Case eFldTy.ePjQ:    O = "ePjQ"
Case eFldTy.eOne:    O = "eOne"
Case eFldTy.eSku:    O = "eSku"
Case eFldTy.eCstVal: O = "eCstVal"
Case eFldTy.eCstRmk: O = "eCstRmk"
Case eFldTy.eChr:    O = "eChr"
Case Else: Er "Enm-eFldTy-{MbrVal} not in valid {MbrVal-List} of {MbrNm-List}", P, "[1 2 3 4 5 6]", "[ePjQ eOne eSku eCstVal eCstRmk eChr]"
End Select
Enm_FldTy_ToStr = O
End Function

Function TDtaChg(Src As TSrc) As TDtaChg()
Dim W1 As PjQCno
Dim O1 As PjQCno
Dim W2 As SkuCno
Dim O2 As SkuCno
Dim W3 As OneCno
Dim O3 As OneCno
Dim T As eFldTy
Dim R&
Erase O
WrkKey = Src.Wrk.KeyDta
WrkSqv = Src.Wrk.Sqv
OrgSqv = Src.Org.Sqv
WrkCno = Src.Wrk.Cno
OrgCno = Src.Org.Cno

Z_Chr
Z_Cst
Z_CstRmk
W1 = WrkCno.PjQ
O1 = WrkCno.PjQ
W2 = WrkCno.Sku
O2 = WrkCno.Sku
W3 = WrkCno.One
O3 = OrgCno.One
T = eFldTy.eSku
For R = 1 To Src.Wrk.UR
    ZOneCell R, eOpt, ePjQ, "RateCHF", W1.RateCHF, O1.RateCHF
    ZOneCell R, eOpt, ePjQ, "RateUSD", W1.RateUSD, O1.RateUSD
    ZOneCell R, eOpt, ePjQ, "RateJPY", W1.RateJPY, O1.RateJPY
    ZOneCell R, eOpt, T, "AssWatchHKD", O2.AssWatchHKD, W2.AssWatchHKD
    ZOneCell R, eOpt, T, "AssWatchUSD", O2.AssWatchUSD, W2.AssWatchUSD
    ZOneCell R, eOpt, T, "SalesmanHKD", O2.SalesmanHKD, W2.SalesmanHKD
    ZOneCell R, eOpt, T, "SalesmanUSD", O2.SalesmanUSD, W2.SalesmanUSD
    ZOneCell R, eOpt, T, "CompleteWatchHKD", O2.CompleteWatchHKD, W2.CompleteWatchHKD
    ZOneCell R, eOpt, T, "CompleteWatchUSD", O2.CompleteWatchUSD, W2.CompleteWatchUSD
    ZOneCell R, eMust, T, "Cost", O2.Cost, W2.Cost
    ZOneCell R, eMust, T, "PotentialQty", O2.PotentialQty, W2.PotentialQty
    ZOneCell R, eOpt, eOne, "ProtCst", O3.ProtCst, W3.ProtCst
    ZOneCell R, eOpt, eOne, "ProtRmk", O3.ProtRmk, W3.ProtRmk
    ZOneCell R, eOpt, eOne, "ToolCst", O3.ToolCst, W3.ToolCst
    ZOneCell R, eOpt, eOne, "ToolRmk", O3.ToolRmk, W3.ToolRmk
Next
TDtaChg = O
End Function

Private Sub ZOneCell(R&, Opt As eOptional, FldTy As eFldTy, FldNm$, OrgCno%, WrkCno%, Optional CostGp$, Optional CostEle$, Optional ChrNm$, Optional ChrCd$)
Dim WrkVal, OrgVal
Dim Key As KeyDta
Dim M As TDtaChg
Dim N&

If Opt = eOpt Then
    If WrkCno = 0 Then Exit Sub
End If
WrkVal = WrkSqv(R, WrkCno)
If OrgCno > 0 Then
    OrgVal = OrgSqv(R, OrgCno)
End If
If WrkVal <> OrgVal Then
    Key = WrkKey(R)
    With M
        .Cno = WrkCno
        .CharName = ChrNm
        .CharCode = ChrCd
        .CostGp = CostGp
        .CostEle = CostEle
        .FldNm = FldNm
        .FldTy = FldTy
        .WrkVal = WrkVal
        .OrgVal = OrgVal
        .Key = Key
    End With
    N = ZSz
    ReDim Preserve O(N)
    O(N) = M
End If
End Sub

Private Function ZSz&()
On Error Resume Next
ZSz = UBound(O) + 1
End Function

Private Sub Z_Chr()
Dim WAy() As ChrCno
Dim OAy() As ChrCno
Dim J%
Dim W As ChrCno
Dim O As ChrCno
Dim Fnd As Boolean
Dim I%
Dim R&

OAy = OrgCno.Chr
WAy = WrkCno.Chr
For J = 0 To UBound(WAy)
    W = WAy(J)
    Fnd = False
    For I = 0 To UBound(OAy)
        O = OAy(I)
        If O.CostGp = W.CostGp Then
        If O.CostEle = W.CostEle Then
        If O.CharName = W.CharName Then
        If O.CharCode = W.CharCode Then
            Fnd = True
            Exit For
        End If
        End If
        End If
        End If
    Next
    If Fnd Then
        For R = 1 To UBound(WrkSqv, 1)
            ZOneCell R, eOpt, eFldTy.eChr, "Char", O.Cno, W.Cno, W.CostGp, W.CostEle, W.CharName, W.CharCode
        Next
    End If
Next
End Sub

Private Sub Z_Cst()
Dim OAy() As CstValCno
Dim WAy() As CstValCno
Dim J%
Dim W As CstValCno
Dim O As CstValCno
Dim Fnd As Boolean
Dim I%
Dim R&
OAy = WrkCno.CstVal
WAy = WrkCno.CstVal
For J = 0 To UBound(WAy)
    W = WAy(J)
    Fnd = False
    For I = 0 To UBound(OAy)
        O = OAy(I)
        If O.CostGp = W.CostGp Then
            If O.CostEle = W.CostEle Then
                Fnd = True
                Exit For
            End If
        End If
    Next
    If Fnd Then
        For R = 1 To UBound(WrkSqv, 1)
            ZOneCell R, eOpt, eFldTy.eCstVal, "Cost", O.Cno, W.Cno, W.CostGp, W.CostEle
        Next
    End If
Next
End Sub

Private Sub Z_CstRmk()
Dim WAy() As CstRmkCno
Dim OAy() As CstRmkCno
Dim J%
Dim W As CstRmkCno
Dim O As CstRmkCno
Dim Fnd As Boolean
Dim I%
Dim R&
WAy = WrkCno.CstRmk
OAy = OrgCno.CstRmk
For J = 0 To UBound(WAy)
    W = WAy(J)
    Fnd = False
    For I = 0 To UBound(OAy)
        O = OAy(I)
        If W.CostEle = O.CostEle Then
        If W.CostGp = O.CostGp Then
            Fnd = True
            Exit For
        End If
        End If
    Next
    If Fnd Then
        For R = 1 To UBound(WrkSqv, 1)
            ZOneCell R, eOpt, eCstRmk, "Cost Rmk", O.Cno, W.Cno, W.CostGp, W.CostEle     '<===
        Next
    End If
Next
End Sub
