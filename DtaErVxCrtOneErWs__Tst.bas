Attribute VB_Name = "DtaErVxCrtOneErWs__Tst"
Option Explicit

Sub Tst()
Dim Key As KeyDta
    Key.Pj = "Pj"
    Key.QDte = Now
    Key.Sku = "Sku"
Dim ChrDef As ChrDef
    ChrDef = ChrDefInf.ChrDef(0)
Dim D As TDtaErOpt
    Dim F() As TCnoDef
        F = Src.Wrk.CnoDef
    DtaErOpt_Push_EmptyChr D, ChrDef, "A1", Key, "OrgVal"
    DtaErOpt_Push_ChrCdNotFnd D, ChrDef, "A1", "ZXXX_Er"
    DtaErOpt_Push_ChrVal D, ChrDef, "A1", "ErVal", "OrgVal", Key
    DtaErOpt_Push_DifHdCell D, "OrgHdVal", "WrkHdVal", "A1"
    DtaErOpt_Push_DifColCnt D, "A1", 2, 3
    DtaErOpt_Push_DifR1Formula D, eCstTotOfMultiEle, Key, 123, "A1", "CostEle", "CostGp", "R1Fomula", "ErFomula", OrgWsNm
    DtaErOpt_Push_DifVal D, "FldNm", "A1", Key, "OrgAdr", "OrgVal", "ErVal"
    DtaErOpt_Push_DupSku D, eOrgWs, Key, "A1", 123
    DtaErOpt_Push_NoOrgRow D, "A1", Key
    DtaErOpt_Push_ValTy D, eWrkWs, Key, F(0), "AA", 1, "Empty"
'---
Dim WrkWs As Worksheet
    Set WrkWs = Src.Wrk.Ws
DtaEr_DoCrt_ErWs D, WrkWs, eV4SelInSep
End Sub
