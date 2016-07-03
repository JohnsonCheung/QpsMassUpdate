Attribute VB_Name = "DtaErOpt"
Option Explicit

Sub DtaErOpt_Push(OA As TDtaErOpt, M As TDtaEr)
Dim N&
N = DtaErOpt_Sz(OA)
ReDim Preserve OA.Ay(N)
OA.Ay(N) = M
End Sub

Sub DtaErOpt_PushAy(OA As TDtaErOpt, A As TDtaErOpt)
Dim J&
For J = 0 To DtaErOpt_Sz(A) - 1
    DtaErOpt_Push OA, A.Ay(J)
Next
End Sub

Sub DtaErOpt_Push_ChrCdNotFnd(OA As TDtaErOpt, ChrDef As ChrDef, Adr$, ErChrCd$)
Dim M As TDtaEr
M.Ty = eChrCdNotFndEr
With M.ChrCdNotFnd.ShwFld
     .Adr = Adr
     .CharName = ChrDef.CharName
     .CostEle = ChrDef.CostEle
     .CostGp = ChrDef.CostGp
End With
With M.ChrCdNotFnd.MsgDta
    .ErChrCd = ErChrCd
End With
DtaErOpt_Push OA, M
End Sub

Sub DtaErOpt_Push_ChrVal(OA As TDtaErOpt, D As ChrDef, Adr$, ErVal$, OrgVal$, Key As KeyDta)
Dim M As TDtaEr
M.Ty = eChrValEr
M.ChrVal.ChrDef = D
With M.ChrVal.ShwFld
    .Adr = Adr
    .CharName = D.CharName
    .CostEle = D.CostEle
    .CostGp = D.CostGp
    .Key = Key
    .OrgVal = OrgVal
    .ErVal = ErVal
End With
DtaErOpt_Push OA, M
End Sub

Sub DtaErOpt_Push_DifColCnt(OA As TDtaErOpt, Adr$, OrgHdColSz%, WrkHdColSz%)
Dim M As TDtaEr
M.Ty = eDifColCntEr
With M.DifColCnt.ShwFld
    .Adr = Adr
'    .CharName = Def.CharName
'    .CostEle = Def.CostEle
'    .CostGp = Def.CostGp
'    .Key = Key
'    .OrgVal = OrgVal
End With
With M.DifColCnt.MsgDta
    .OrgHdColSz = OrgHdColSz
    .WrkHdColSz = WrkHdColSz
End With
DtaErOpt_Push OA, M
End Sub

Sub DtaErOpt_Push_DifHdCell(OA As TDtaErOpt, OrgHdVal$, WrkHdVal$, Adr$)
Dim M As TDtaEr
M.Ty = eDifHdCellEr
With M.DifHdCell.ShwFld
    .Adr = Adr
End With
With M.DifHdCell.MsgDta
    .OrgHdVal = OrgHdVal
    .WrkHdVal = WrkHdVal
End With
DtaErOpt_Push OA, M
End Sub

Sub DtaErOpt_Push_DifR1Formula(OA As TDtaErOpt, ChkCol As eR1FormulaChkCol, Key As KeyDta, RFirst&, Adr$, CostEle$, CostGp$, R1Formula$, ErFormula$, Ws$)
Dim Shw As ShwFld_DifR1Formula
    Shw.Adr = Adr
    Shw.CostEle = CostEle
    Shw.CostGp = CostGp
    Shw.Key = Key
    Shw.Ws = Ws
    Shw.FldNm = Enm.R1FormulaChkColStr(ChkCol)
Dim Msg As MsgDta_DifR1Formula
    Msg.R1Formula = R1Formula
    Msg.ErFormula = ErFormula
Dim MM As DE_DifR1Formula
    MM.MsgDta = Msg
    MM.ShwFld = Shw
Dim M As TDtaEr
    M.Ty = eDifR1FormulaEr
    M.DifR1Formula = MM
DtaErOpt_Push OA, M
End Sub

Sub DtaErOpt_Push_DifVal(OA As TDtaErOpt, FldNm$, Adr$, Key As KeyDta, OrgAdr$, OrgVal$, ErVal)
Dim Shw As ShwFld_DifVal
    With Shw
        .Adr = Adr
        .FldNm = FldNm
        .Key = Key
        .ErVal = ErVal
    End With
    
Dim Msg As MsgDta_DifVal
    With Msg
        .OrgAdr = OrgAdr
        .OrgVal = OrgVal
        .FldNm = FldNm
    End With

Dim M As TDtaEr
    M.Ty = eDifValEr
    M.DifVal.MsgDta = Msg
    M.DifVal.ShwFld = Shw
DtaErOpt_Push OA, M
End Sub

Sub DtaErOpt_Push_DupSku(OA As TDtaErOpt, WhichWs As eWhichWs, Key As KeyDta, Adr$, FirstRno_WithDupSku&)
Dim M As TDtaEr
M.Ty = eDupSkuEr
With M.DupSku.ShwFld
    .Ws = Enm.WhichWsNm(WhichWs)
     .Adr = Adr
     .Key = Key
     .FldNm = "Pj+Sku+QDte"
End With
M.DupSku.MsgDta.FirstRno_WithDupSku = FirstRno_WithDupSku
DtaErOpt_Push OA, M
End Sub

Sub DtaErOpt_Push_EmptyChr(OA As TDtaErOpt, D As ChrDef, Adr$, Key As KeyDta, OrgVal)
Dim M As TDtaEr
M.Ty = eChrEmptyEr
M.ChrEmpty.ChrDef = D
With M.ChrEmpty.ShwFld
    .Adr = Adr
    .CharName = D.CharName
    .CostEle = D.CostEle
    .CostGp = D.CostGp
    .Key = Key
    .OrgVal = OrgVal
End With
DtaErOpt_Push OA, M
End Sub

Sub DtaErOpt_Push_NoOrgRow(OA As TDtaErOpt, Adr$, Key As KeyDta)
Dim Shw As ShwFld_NoOrgRow
With Shw
    .Adr = Adr
    .Key = Key
    .FldNm = "Pj+Sku+QDte"
End With
Dim M As TDtaEr
M.Ty = eNoOrgRowEr
M.NoOrgRow.ShwFld = Shw
DtaErOpt_Push OA, M
End Sub

Sub DtaErOpt_Push_ValTy(OA As TDtaErOpt, WhichWs As eWhichWs, Key As KeyDta, F As TCnoDef, ErVal, R&, ActTy$)
Dim Shw As ShwFld_ValTy
Dim Msg As MsgDta_ValTy
With Msg
    .ErVal = ErVal
    .ExpDtaTy = Enm_ValTy_ToStr(F.ValTy)
    .ActTy = ActTy
End With
With Shw
    .Key = Key
    .ErVal = ErVal
    .Adr = Fct.SrcSqvAdr(R, F.Cno)
    .Ws = Enm.WhichWsNm(WhichWs)
    .FldNm = F.Nm
    .CostEle = F.CostEle
    .CostGp = F.CostGp
    .CharName = F.CharName
End With
Dim M As TDtaEr
    M.Ty = eDtaErTy.eValTyEr
    M.ValTy.ShwFld = Shw
    M.ValTy.MsgDta = Msg
DtaErOpt_Push OA, M
End Sub

Property Get DtaErOpt_Sz&(A As TDtaErOpt)
If A.Some Then
    DtaErOpt_Sz = UBound(A.Ay) + 1
End If
End Property
