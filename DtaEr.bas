Attribute VB_Name = "DtaEr"
Option Explicit
Enum eR1FormulaChkCol     ' What column will do this checking
    eCstTotOfMultiEle = 1 ' The formula of the column of total-cst of multi-element should all be the same
    eSkuCost = 2
End Enum

Type MsgDta_DifR1Formula
    ErFormula As String
    R1Formula As String
End Type
Type MsgDta_DifVal
    FldNm As String
    OrgAdr As String
    OrgVal As String
End Type
Type MsgDta_ValTy
    ErVal As Variant
    ExpDtaTy As String
    ActTy As String
End Type
Type MsgDta_DupSku
    FirstRno_WithDupSku As Long
End Type
Type MsgDta_DifHdCell
    WrkHdVal As String
    OrgHdVal As String
End Type
Type MsgDta_DifColCnt
    WrkHdColSz As Integer
    OrgHdColSz As Integer
End Type
Type MsgDta_ChrCdNotFnd
    ErChrCd As String
End Type
'---
Type TErTxt
    Msg As String
    Sht As String
End Type
'================================================================
'=== ShwFld_{Er} ==============
Type ShwFld_DupSku
    Key As KeyDta
    Ws As String
    Adr As String   ' Adr of Sku with Duplication
    FldNm As String
End Type
Type ShwFld_ValTy
    Ws As String
    Adr As String
    Key As KeyDta
    CostGp As String
    CostEle As String
    CharName As String
    FldNm As String
    ErVal As Variant
End Type
Type ShwFld_ChrCdNotFnd
    Adr As String
    FldNm As String  ' Use FldNm as CharCode, because on the error worksheet, no CharCOde is shown
    CharName As String
    CostEle As String
    CostGp As String
End Type

Type ShwFld_ChrVal
    Ws As String
    Adr As String
    ErVal As String
    OrgVal As Variant ' Value from OrgWs. In ErWs, it is formatted as Left-Aligned.
                      ' keep the value in it orginal type so that it can be used as restore
    CharName As String
    CostEle As String
    CostGp As String
    Key As KeyDta
End Type
Type ShwFld_ChrEmpty
    Ws As String
    Adr As String
    OrgVal As Variant ' Value from OrgWs. In ErWs, it is formatted as Left-Aligned.
                      ' keep the value in it orginal type so that it can be used as restore
    CharName As String
    CostEle As String
    CostGp As String
    Key As KeyDta
End Type
Type ShwFld_DifR1Formula
     Adr As String
     CostEle As String
     CostGp As String
     Key As KeyDta
     FldNm As String
     Ws As String
End Type
Type ShwFld_DifVal
    ErVal As String
    Adr As String
    Key As KeyDta
    FldNm As String
End Type
Type ShwFld_DifHdCell
    Adr As String   ' Adr of the cell with dif-value
End Type
Type ShwFld_DifColCnt
    Adr As String   ' ="A" & LastCol
    WsNmWhichIsLargerNoOfCol As String    ' Worksheet have the larger column
End Type
Type ShwFld_NoOrgRow
    Adr As String
    Key As KeyDta
    FldNm As String
End Type
'=== DE_{Er} ==============
Type DE_DupSku ' There 2 or more row with same Pj+Sku+QDte
    ShwFld As ShwFld_DupSku
    MsgDta As MsgDta_DupSku
End Type

Type DE_ChrVal    ' DE_ means DeAy_
    ShwFld As ShwFld_ChrVal
    ChrDef As ChrDef
End Type
Type DE_ChrEmpty    ' DE_ means DeAy_
    ShwFld As ShwFld_ChrEmpty
    ChrDef As ChrDef
End Type
Type DE_DifR1Formula    ' Formula is dif from Row-1
    ShwFld As ShwFld_DifR1Formula
    MsgDta As MsgDta_DifR1Formula
End Type
Enum eDifValFld
    eBrd = 1
    eVnd = 2
End Enum
Type DE_DifVal ' The XXX(Brand|Supplier) is Dif is changed in Ws-Wrk
    ShwFld As ShwFld_DifVal
    MsgDta As MsgDta_DifVal
End Type

Type DE_DifHdCell
    ShwFld As ShwFld_DifHdCell
    MsgDta As MsgDta_DifHdCell
End Type
Type DE_DifColCnt
    ShwFld As ShwFld_DifColCnt
    MsgDta As MsgDta_DifColCnt
End Type
Type DE_NoOrgRow
    ShwFld As ShwFld_NoOrgRow
    FldNm As String
End Type
Type DE_ValTy
    ShwFld As ShwFld_ValTy
    MsgDta As MsgDta_ValTy
End Type
Type DE_ChrCdNotFnd
    ShwFld As ShwFld_ChrCdNotFnd
    MsgDta As MsgDta_ChrCdNotFnd
End Type

Enum eDtaErTy
    eChrCdNotFndEr = 1
    eChrValEr = 2
    eChrEmptyEr = 3  ' An empty CharValName is entered, but it is a must field
    eDifColCntEr = 4
    eDifHdCellEr = 5
    eDifR1FormulaEr = 7
    eDifValEr = 8
    eDupSkuEr = 9
    eNoOrgRowEr = 10
    eValTyEr = 11
End Enum
'=====
Type TDtaEr
    Ty As eDtaErTy
    ChrCdNotFnd  As DE_ChrCdNotFnd
    ChrVal       As DE_ChrVal
    ChrEmpty     As DE_ChrEmpty
    DifR1Formula As DE_DifR1Formula
    DifHdCell    As DE_DifHdCell
    DifColCnt    As DE_DifColCnt
    DifVal       As DE_DifVal
    DupSku       As DE_DupSku
    NoOrgRow     As DE_NoOrgRow
    ValTy        As DE_ValTy
End Type
'======================
Type TDtaErOpt
    Ay() As TDtaEr
    Some As Boolean
End Type

Private O As TDtaErOpt
Private Src As TSrc

Function TDtaEr(P As TSrc) As TDtaErOpt
Erase O.Ay
Src = P
ZNoOrgRow
ZChrVal
ZDifHdCell
ZDifColCnt
ZDifR1Formula
ZDifVal
ZValTy
ZDupSku
ZChrCdNotFnd
ZEmptyChr
O.Some = Z_Sz > 0
TDtaEr = O
End Function

Private Sub ZChrCdNotFnd()
Dim VdtChrCdAy$()
    VdtChrCdAy = ChrDefInf.VdtChrCdAy
Dim WsI%
For WsI = 0 To 1
    Dim WsInf As TWsInf
        Select Case WsI
        Case 0: WsInf = Src.Wrk
        Case 1: WsInf = Src.Org
        Case Else: Er "Invalid {WsI}", WsI
        End Select
    Dim WsChrCno() As ChrCno:   WsChrCno = WsInf.Cno.Chr
    Dim WsHdLinChr%:          WsHdLinChr = WsInf.HdLinChr
    Dim ColN%:                      ColN = WsInf.Cno.NChr
    Dim ColI%
    For ColI = 0 To ColN - 1
        Dim Col_Cno%
            Col_Cno = WsChrCno(ColI).Cno
        Dim ColChrCd$
            ColChrCd = WsInf.HdSqv(WsHdLinChr, Col_Cno)
        Dim Col_IsChrCd_NotFnd As Boolean
            Col_IsChrCd_NotFnd = Not Ay_Has(VdtChrCdAy, ColChrCd)
        If Not Col_IsChrCd_NotFnd Then GoTo Nxt_Col     '<=== No Error goto Nxt_Chr
        
        Dim D As ChrDef
            D = ChrDefInf.ChrCd_ChrDef(ColChrCd)
        
        Dim Msg As MsgDta_ChrCdNotFnd
            Msg.ErChrCd = ColChrCd
        Dim Shw As ShwFld_ChrCdNotFnd
            Shw.Adr = Fct.SrcSqvAdr(0, Col_Cno)
            Shw.CharName = D.CharName
            Shw.CostEle = D.CostEle
            Shw.CostGp = D.CostGp
            Shw.FldNm = "Char"
        Dim MM As DE_ChrCdNotFnd
            MM.MsgDta = Msg
            MM.ShwFld = Shw
        Dim M As TDtaEr
            M.Ty = eChrCdNotFndEr
            M.ChrCdNotFnd = MM
                    
        Z_Push M
Nxt_Col:
    Next ColI
Next WsI
End Sub

Private Sub ZChrCdNotFnd__Tst()
Src = SrcInf.Src
ZChrCdNotFnd
Dim Act As TDtaErOpt
Act = O
Stop
End Sub

Private Sub ZChrVal()
Dim WsSqv
Dim WsI%
Dim WsInf As TWsInf
Dim WsChrCno() As ChrCno
Dim WsKey() As KeyDta
Dim ChrN%
Dim ChrI%
Dim ChrCd$
Dim ChrCno%
Dim ChrDef As ChrDef
Dim ChrIsNeedInList As Boolean
Dim ChrIsMulti As Boolean
Dim R&
Dim V$
Dim Er_V$()
Dim VV$
Dim Er_I
Dim M As TDtaEr
Dim Shw As ShwFld_ChrVal

For WsI = 2 To 2
    Select Case WsI
    Case 1: WsInf = Src.Org
    Case 2: WsInf = Src.Wrk
    End Select
        
    WsSqv = WsInf.Sqv
    With WsInf.Cno
        ChrN = .NChr
        WsChrCno = .Chr
    End With
    WsKey = WsInf.KeyDta
        
    For ChrI = 0 To ChrN - 1
        With WsChrCno(ChrI)
            ChrCd = .CharCode
            ChrCno = .Cno
        End With
    
        ChrDef = ChrDefInf.ChrCd_ChrDef(ChrCd)
    
        ChrIsNeedInList = ChrDef.IsNeedInList
            
        If Not ChrIsNeedInList Then GoTo ChrNxt
        
        ChrIsMulti = ChrDef.IsMulti
        
        For R = 1 To UBound(WsSqv, 1)
            If VarType(WsSqv(R, ChrCno)) = vbError Then GoTo R_Nxt

            V = Trim(WsSqv(R, ChrCno))
            
            If V = "" Then GoTo R_Nxt     'Empty cell will be check in EmptyChrEr
            
            Erase Er_V
            If ChrIsMulti Then
                Dim V_Ay$()
                    V_Ay = Split(V, vbLf)
                
                Dim V_I%
                For V_I = 0 To UB(V_Ay)
                    VV = Trim(V_Ay(V_I))
                        
                    If Not ChrDef.Dic_OfValNm_ToValCd.Exists(VV) Then
                        Push Er_V, VV
                    End If
                Next
            Else
                If Not ChrDef.Dic_OfValNm_ToValCd.Exists(V) Then
                    Push Er_V, V
                End If
            End If
                
            For Er_I = 0 To UB(Er_V)
                With Shw
                    .Ws = IIf(WsI = 1, OrgWsNm, WrkWsNm)
                    .CharName = ChrDef.CharName
                    .Key = WsKey(R)
                    .Adr = Fct.SrcSqvAdr(R, ChrCno)
                    .CostEle = ChrDef.CostEle
                    .CostGp = ChrDef.CostGp
                    .ErVal = V
                End With
                M.Ty = eDtaErTy.eChrValEr
                M.ChrVal.ChrDef = ChrDef
                M.ChrVal.ShwFld = Shw
                Z_Push M                    '<=== Push Er
            Next
R_Nxt:
        Next
ChrNxt:
    Next
Next
End Sub

Private Sub ZChrVal__Tst()
Src = SrcInf.Src
ZChrVal
Dim Act As TDtaErOpt
Act = O
Stop
End Sub

Private Sub ZDifColCnt()
Dim OrgC&
Dim WrkC&
    WrkC = UBound(Src.Wrk.HdSqv, 2)
    OrgC = UBound(Src.Org.HdSqv, 2)

If OrgC = WrkC Then Exit Sub    '<=== No Error

Dim Msg As MsgDta_DifColCnt
    Msg.OrgHdColSz = OrgC
    Msg.WrkHdColSz = WrkC

Dim Shw As ShwFld_DifColCnt
    If OrgC > WrkC Then
        Shw.Adr = Ws_Adr(yWsMassUpd, 6, OrgC)
        Shw.WsNmWhichIsLargerNoOfCol = OrgWsNm
    Else
        Shw.Adr = Ws_Adr(yWsMassUpd, 6, WrkC)
        Shw.WsNmWhichIsLargerNoOfCol = WrkWsNm
    End If
    
Dim M As TDtaEr
    M.Ty = eDifColCntEr
    M.DifColCnt.MsgDta = Msg
    M.DifColCnt.ShwFld = Shw

Z_Push M '<====
End Sub

Private Sub ZDifHdCell()
Dim OrgH, WrkH
    OrgH = Src.Org.HdSqv
    WrkH = Src.Wrk.HdSqv
    
If UBound(OrgH, 2) <> UBound(WrkH, 2) Then Exit Sub

    Dim UR%, UC%
        UR = UBound(OrgH, 1)
        UC = UBound(OrgH, 2)
    Dim R&, C%
    For R = 1 To UR
        For C = 1 To UC
            Dim V1, V2
                V1 = OrgH(R, C)
                V2 = WrkH(R, C)
            If V1 = V2 Then GoTo Nxt_Cell '<=== Two cells are equal, goto Nxt_Cell
            
            Dim Shw As ShwFld_DifHdCell
            Dim Msg As MsgDta_DifHdCell
                Msg.OrgHdVal = V1
                Msg.WrkHdVal = V2
                Shw.Adr = Ws_Adr(yWsMassUpd, R, C)
                
            Dim M As TDtaEr
                M.Ty = eDtaErTy.eDifHdCellEr
                M.DifHdCell.MsgDta = Msg
                M.DifHdCell.ShwFld = Shw
            Z_Push M
Nxt_Cell:
        Next
    Next
End Sub

Private Sub ZDifR1Formula()
Dim Ws_I%
Dim Ws_Inf As TWsInf
Dim Ws As Worksheet
Dim Ws_DtaRge As Range
Dim Ws_Cno As TCno
Dim Ws_NCstTot%
Dim Ws_KeyDta() As KeyDta
Dim Col_I%
Dim Col_Ty As eR1FormulaChkCol
Dim Col_ChrCno As ChrCno
Dim Col_FldNm$
Dim Col_Cno%
Dim Col_Rge As Range
Dim Formula_R1$
Dim Rno&
Dim Formula_Cur$
Dim Msg As MsgDta_DifR1Formula
Dim Shw As ShwFld_DifR1Formula
Dim MM As DE_DifR1Formula
Dim M As TDtaEr

For Ws_I = 0 To 1   '<==== Loop Ws
    Select Case Ws_I
    Case 0: Ws_Inf = Src.Org
    Case 1: Ws_Inf = Src.Wrk
    End Select
    Set Ws = Ws_Inf.Ws
    Set Ws_DtaRge = Ws.ListObjects(1).DataBodyRange
    Ws_Cno = Ws_Inf.Cno
    Ws_NCstTot = Ws_Cno.NCstTot
    Ws_KeyDta = Ws_Inf.KeyDta
        
    For Col_I = 0 To Ws_NCstTot             '<=== Loop Formula Columns
        If Col_I = Ws_NCstTot Then
            Col_Ty = eSkuCost
        Else
            Col_Ty = eCstTotOfMultiEle
        End If
            
        If Col_I <> Ws_NCstTot Then
            Col_ChrCno = Ws_Cno.Chr(Col_I)
        End If
            
        Select Case Col_Ty
        Case eR1FormulaChkCol.eCstTotOfMultiEle: Col_FldNm = "Ele Total"
        Case eR1FormulaChkCol.eSkuCost:          Col_FldNm = "Sku Total"
        End Select
        
        Select Case Col_Ty
        Case eR1FormulaChkCol.eCstTotOfMultiEle: Col_Cno = Ws_Cno.CstTot(Col_I).Cno  '<== Col_No
        Case eR1FormulaChkCol.eSkuCost:          Col_Cno = Ws_Cno.Sku.Cost           '<== Col_No
        End Select
            
        Set Col_Rge = Rge_C(Ws_DtaRge, Col_Cno)
        
        Formula_R1 = Rge_RC(Col_Rge, 1, 1).Formula

        For Rno = 2 To Col_Rge.Rows.Count     '<===============================Loop Rows
            Formula_Cur = Rge_RC(Col_Rge, Rno, 1).Formula
            If Formula_Cur = Formula_R1 Then GoTo Nxt_Row                  '<==== No Error, goto Nxt_Row
            
            Msg.ErFormula = Formula_Cur
            Msg.R1Formula = Formula_R1
            
            Shw.FldNm = Col_FldNm
            Shw.CostEle = Col_ChrCno.CostEle
            Shw.CostGp = Col_ChrCno.CostGp
            
            Shw.Adr = Ws_DtaRge(Rno, Col_Cno).Address
            Shw.Key = Ws_KeyDta(Rno)
            Shw.Ws = Ws.Name
            
            MM.MsgDta = Msg
            MM.ShwFld = Shw
            M.Ty = eDifR1FormulaEr
            M.DifR1Formula = MM
            
            Z_Push M              '<====
Nxt_Row:
        Next
    Next
Next
End Sub

Private Sub ZDifR1Formula__Tst()
Src = SrcInf.Src
ZDifR1Formula
Dim Act As TDtaErOpt
Act = O
End Sub

Private Sub ZDifVal()
Dim OrgPkDic As Dictionary
Dim NRno&
Dim WrkKey() As KeyDta
Dim WrkSqv
Dim OrgSqv
Dim WrkBrandCno%, WrkSupplierCno%
Dim OrgBrandCno%, OrgSupplierCno%

NRno = UBound(Src.Wrk.Sqv, 1)
Dim WInf As TWsInf
Dim OInf As TWsInf
    WInf = Src.Wrk
    OInf = Src.Org

WrkKey = WInf.KeyDta

WrkSqv = WInf.Sqv
OrgSqv = OInf.Sqv

WrkBrandCno = WInf.Cno.PjQ.Brand
OrgBrandCno = OInf.Cno.PjQ.Brand
WrkSupplierCno = WInf.Cno.PjQ.Supplier
OrgSupplierCno = OInf.Cno.PjQ.Supplier
Set OrgPkDic = OInf.PkDic

Dim IFld%
For IFld = 1 To 2
    Dim WCno%
    Dim OCno%
        Select Case IFld
        Case 1: OCno = OrgBrandCno:    WCno = WrkBrandCno
        Case 2: OCno = OrgSupplierCno: WCno = WrkSupplierCno
        Case Else: Er "{IFld} error.  Should be 1 or 2", IFld
        End Select
    Dim R&
    For R = 1 To NRno
        Dim KeyStr$
            KeyStr = WrkKey(R).KeyStr
        Dim NoOrgRow As Boolean
            NoOrgRow = Not OrgPkDic.Exists(KeyStr)
    
        If NoOrgRow Then GoTo Nxt_Row '<=== Skip, if there is no org row
        
        Dim WrkVal
        Dim OrgVal
            WrkVal = WrkSqv(R, WCno)
            OrgVal = OrgSqv(R, OCno)
            
        If WrkVal = OrgVal Then GoTo Nxt_Row '<=== No Err, goto next row
            
        Dim Shw As ShwFld_DifVal
            Stop
        Dim Msg As MsgDta_DifVal
            Stop
        Dim MM As DE_DifVal
            MM.ShwFld = Shw
            MM.MsgDta = Msg
        Dim M As TDtaEr
        With M
            M.Ty = eDifValEr
            M.DifVal = MM
        End With
        Z_Push M    '<==== Push
        
Nxt_Row:
    Next R
Next IFld
End Sub

Private Sub ZDupSku()

Dim Wrk As TWsInf
    Wrk = Src.Wrk

Dim Org As TWsInf
    Org = Src.Org
    
Dim IWs%
For IWs = 1 To 2
    Dim WsKey() As KeyDta
    Dim WsSqv
    Dim WsNm$
    Dim WsSkuCno%
        Select Case IWs
        Case 1: WsKey = Org.KeyDta: WsSqv = Org.Sqv: WsNm = OrgWsNm: WsSkuCno = Org.Cno.Key.Sku
        Case 2: WsKey = Wrk.KeyDta: WsSqv = Wrk.Sqv: WsNm = WrkWsNm: WsSkuCno = Wrk.Cno.Key.Sku
        End Select
        
    Dim WsDic As New Dictionary
        WsDic.CompareMode = TextCompare
        WsDic.RemoveAll
        
    Dim R&
    For R = 1 To UBound(WsSqv, 1)
        Dim Key As KeyDta
            Key = WsKey(R)
        
        Dim Sku$
            Sku = Key.Sku
            
        If Not WsDic.Exists(Sku) Then
            WsDic.Add Sku, R
            GoTo Nxt_Row        '<=== No Dup, goto Nxt_Row
        End If
        
        Dim Msg As MsgDta_DupSku
            With Msg
                .FirstRno_WithDupSku = WsDic(Sku)
            End With
        
        Dim Shw As ShwFld_DupSku
            With Shw
                .Adr = Fct.SrcSqvAdr(Key.Rno, WsSkuCno)
                .Ws = WsNm
                .FldNm = "Sku"
                .Key = Key
            End With
        
        Dim M As TDtaEr
            M.Ty = eDupSkuEr
            M.DupSku.MsgDta = Msg
            M.DupSku.ShwFld = Shw
        Z_Push M        '<=== Push to O
Nxt_Row:
    Next
Next
End Sub

Private Sub ZDupSku__Tst()
Src = SrcInf.Src
ZDupSku
Dim Act As TDtaErOpt
Act = O
Stop
End Sub

Private Sub ZEmptyChr()
Dim WsI%
Dim WsInf As TWsInf
Dim WsSqv
Dim WsChrCno() As ChrCno
Dim WsKey() As KeyDta
Dim Chr_N%
Dim Chr_I%
Dim ChrCd$
Dim Cno%
Dim ChrDef As ChrDef
Dim R&
Dim V$
Dim Key As KeyDta
Dim Shw As ShwFld_ChrEmpty
Dim M As TDtaEr

For WsI = 2 To 2
    Select Case WsI
    Case 1: WsInf = Src.Org
    Case 2: WsInf = Src.Wrk
    End Select
    WsSqv = WsInf.Sqv
    With WsInf.Cno
        Chr_N = .NChr
        WsChrCno = .Chr
    End With

    WsKey = WsInf.KeyDta
    For Chr_I = 0 To Chr_N - 1
        With WsChrCno(Chr_I)
            ChrCd = .CharCode
            Cno = .Cno
        End With

        ChrDef = ChrDefInf.ChrCd_ChrDef(ChrCd)
            
        If Not ChrDef.IsNeedInList Then GoTo Nxt_Col
        If Not ChrDef.IsMust Then GoTo Nxt_Col
    
        For R = 1 To UBound(WsSqv, 1)
            V = Trim(WsSqv(R, Cno))
        
            If V <> "" Then GoTo Nxt_Cell       'Empty cell will be check in EmptyChrEr
            Key = WsKey(R)
            
            If IsNothing(ChrDef.Dic_OfValNm_ToValCd) Then Stop
            With Shw
                .Ws = IIf(WsI = 1, OrgWsNm, WrkWsNm)
                .CharName = ChrDef.CharName
                .Key = Key
                .Adr = Fct.SrcSqvAdr(R, Cno)
                .CostEle = ChrDef.CostEle
                .CostGp = ChrDef.CostGp
            End With
            M.Ty = eDtaErTy.eChrEmptyEr
            M.ChrEmpty.ChrDef = ChrDef
            M.ChrEmpty.ShwFld = Shw
            Z_Push M
Nxt_Cell:
        Next
Nxt_Col:
    Next
Next
End Sub

Private Sub ZEmptyChr__Tst()
Src = SrcInf.Src
ZEmptyChr
Dim Act As TDtaErOpt
Act = O
Stop
End Sub

Private Sub ZNoOrgRow()
Dim KeyDta() As KeyDta
   KeyDta = Src.Wrk.KeyDta
   
Dim UR&
    UR = Src.Wrk.UR

Dim SkuCno%
    SkuCno = Src.Wrk.Cno.Key.Sku

Dim OrgPkDic As Dictionary
    Set OrgPkDic = Src.Org.PkDic
    
'For each KeyDta in WrkWs, it must be found in OrgPkDic, otherwise, it is an error
    Dim R&
    For R = 1 To UR
        Dim Key As KeyDta
            Key = KeyDta(R)
            
        Dim KeyStr$
            KeyStr = Key.KeyStr
            
        If OrgPkDic.Exists(KeyStr) Then GoTo Nxt_Row
        
        Dim Adr$
            Adr = Fct.SrcSqvAdr(R, SkuCno)
        Dim Shw As ShwFld_NoOrgRow
            With Shw
                .Adr = Adr
                .FldNm = "Sku"
                .Key = Key
            End With
        
        Dim M As TDtaEr
            M.Ty = eDtaErTy.eNoOrgRowEr
            M.NoOrgRow.FldNm = "Sku"
            M.NoOrgRow.ShwFld = Shw
        Z_Push M                    '<=== Push ErNxt:
Nxt_Row:
    Next
End Sub

Private Sub ZNoOrgRow__Tst()
Src = SrcInf.Src
ZNoOrgRow
Dim Act As TDtaErOpt
Act = O
Stop
End Sub

Private Sub ZValTy()
Dim Ws_I%
For Ws_I = 0 To 1
    Dim WsInf As TWsInf
        Select Case Ws_I
        Case 0: WsInf = Src.Wrk
        Case 1: WsInf = Src.Org
        End Select
    
    Dim WsSqv
    Dim WsCno As TCno
    Dim WsChrCno() As ChrCno
        WsCno = WsInf.Cno
        WsChrCno = WsCno.Chr
        WsSqv = WsInf.Sqv
    Dim Rno_N&
        Rno_N = UBound(WsSqv, 1)
    Dim KeyDta() As KeyDta
        KeyDta = WsInf.KeyDta
        
    Dim ValTy_Ay() As TCnoDef
        ValTy_Ay = TCnoDef(WsCno)
    Dim Col_I%
    For Col_I = 0 To UBound(ValTy_Ay)
        Dim Col_Ty As eValTy
        Dim Col_Nm$
        Dim Col_Cno%
            With ValTy_Ay(Col_I)
                Col_Ty = .ValTy
                Col_Cno = .Cno
                Col_Nm = .Nm
            End With

        Dim ChrDef As ChrDef
        Dim IsOptCol_Missing As Boolean
            IsOptCol_Missing = False
            If Col_Cno = 0 Then
                Select Case Col_Ty
                Case eDteOpt, ePosOpt, eNbrOpt, eStrOpt:  IsOptCol_Missing = True
                End Select
            End If

        If IsOptCol_Missing Then GoTo Col_Nxt
        Dim Rno_I&
        For Rno_I = 1 To Rno_N
            Dim V
                V = WsSqv(Rno_I, Col_Cno)

            Dim IsOptAndEmpty As Boolean
                IsOptAndEmpty = False
                If IsEmpty(V) Then
                    Select Case Col_Ty
                    Case eStrOpt, eDteOpt, eNbrOpt, ePosOpt: IsOptAndEmpty = True
                    End Select
                End If
            If IsOptAndEmpty Then
                GoTo Rno_Nxt
            End If
            
            Dim IsEr As Boolean
            Dim ActTy$
                ActTy = "?"
                IsEr = False
                If IsEmpty(V) Then
                    IsEr = True
                    ActTy = "Empty"
                Else
                    Dim VbTy As VbVarType
                        VbTy = VarType(V)
                        
                    Select Case Col_Ty
                    Case eStr, eStrOpt: If VbTy <> vbString Then ActTy = TypeName(V): IsEr = True
                    Case eDte, eDteOpt: If VbTy <> vbDate Then ActTy = TypeName(V): IsEr = True
                    Case eNbr, eNbrOpt: If VbTy <> vbDouble Then ActTy = TypeName(V): IsEr = True
                    Case ePos, ePosOpt
                        If VbTy <> vbDouble Then
                            ActTy = TypeName(V)
                            IsEr = True
                        Else
                            If V <= 0 Then
                                ActTy = "-ve Double"
                                IsEr = True
                            End If
                        End If
                    Case Else: Er "{ValTy} error", Col_Ty
                    End Select
                End If
            
            If Not IsEr Then
                GoTo Rno_Nxt
            End If
            
            Dim Msg As MsgDta_ValTy
                With Msg
                    .ActTy = ActTy
                    .ErVal = V
                    .ExpDtaTy = Col_Ty
                End With
            Dim Shw As ShwFld_ValTy
                With Shw
                    .CharName = ChrDef.CharName
                    .CostEle = ChrDef.CostEle
                    .CostGp = ChrDef.CostGp
                    .ErVal = V
                    .FldNm = Col_Nm
                    .Key = KeyDta(Rno_I)
                End With
            Dim MM As DE_ValTy
                MM.MsgDta = Msg
                MM.ShwFld = Shw
            Dim M As TDtaEr
                M.Ty = eValTyEr
                M.ValTy = MM
            Z_Push M
Rno_Nxt:
        Next Rno_I
Col_Nxt:
    Next Col_I
Next Ws_I
End Sub

Private Sub ZValTy__Tst()
Src = SrcInf.Src
ZValTy
Dim Act As TDtaErOpt
Act = O
Stop
End Sub

Private Sub Z_Push(M As TDtaEr)
Dim N&
    N = Z_Sz
ReDim Preserve O.Ay(N)
    O.Ay(N) = M
    O.Some = True
End Sub

Private Property Get Z_Sz%()
On Error Resume Next
Z_Sz = UBound(O.Ay) + 1
End Property
