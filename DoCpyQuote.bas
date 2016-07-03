Attribute VB_Name = "DoCpyQuote"
Option Explicit
Private Db As Database
Private NewQDte As Date
Const TblNmLvs_Pj_ = "ProjQ ProjOneTimeCost"
Const TblNmLvs_Sku = "Sku SkuCostEle SkuCostChr"

Sub Do_CpyQuote(P_Db As Database, KeyDta() As KeyDta, P_NewQDte As Date)
Dim PjKey() As PjKey
Dim P As PjKey
Dim K As KeyDta
Dim SkuWhere$
Dim PjWhere$
Dim J&
Dim Pj$, QDte As Date, Sku$

PjKey = ZPjKey(KeyDta)
NewQDte = P_NewQDte
Set Db = P_Db

ZAssert_PjKeyCannotDup PjKey
ZAssert_NewQDte_ShouldBeNewest NewQDte

For J = 0 To UBound(PjKey)
    P = PjKey(J)
    PjWhere = Fmt_QQ("ProjNo='?' and QuoteDate=#?#", P.Pj, P.QDte)
    ZCpy_TblPjQ PjWhere
    ZCpy_TblOne PjWhere
Next

For J = 0 To UBound(KeyDta)
    K = KeyDta(J)
    SkuWhere = Fmt_QQ("ProjNo='?' and QuoteDate=#?# and Sku='?'", K.Pj, K.QDte, K.Sku)
    ZCpy_TblSku SkuWhere
    ZCpy_TblSkuCostEle SkuWhere   ' ZCpy_TblSkuCostEle Must before ZCpy_TblSkuCostChr,
                                  ' because ZCpy_TblSkuCostChr must have ZCpy_TblSkuCostEle exist first
    ZCpy_TblSkuCostChr SkuWhere
Next
End Sub

Sub Tst()
Dim ToWherePj$, ToWhereSku$
Dim KeyDta() As KeyDta
Dim NewQDte As Date
Dim RecCntAy_Fm&()

Set Db = Dao.DBEngine.OpenDatabase(Fs_Fb)

ZTst_FndTstDta NewQDte, ToWherePj, ToWhereSku, KeyDta, RecCntAy_Fm
'-----------------
ZTst_Assert_RecCnt ToWherePj, ToWhereSku, LngAy(0, 0, 0, 0, 0)
Do_CpyQuote Db, KeyDta, NewQDte                                 '<=== Do_Cpy
ZTst_Assert_RecCnt ToWherePj, ToWhereSku, RecCntAy_Fm
Db.Close
Set Db = Nothing
End Sub

Sub ZDltPj(P As PjKey)
Dim WherePj$
    WherePj = Fmt("ProjNo='{0}' and QuoteDate=#{1}#", P.Pj, P.QDte)
Z_RunDlt WherePj, "SkuCostChr SkuCostEle Sku ProjOneTimeCost ProjQ" ' Must in this dependence order
End Sub

Property Get ZPjKey(Ay() As KeyDta) As PjKey()
'Return uniq-PjKeyAy
Dim A$()            ' PjKeyStr Ay
Dim J&
Dim O() As PjKey
Dim PjKeyStr$
Dim Pj$, QDte As Date
Dim U&
Dim B$()
Dim M As PjKey
For J = 0 To UBound(Ay)
    Pj = Ay(J).Pj
    QDte = Ay(J).QDte
    PjKeyStr = Pj & "|" & QDte
    Push_NoDup A, PjKeyStr  '<===
Next

U = UB(A)
ReDim O(U)
For J = 0 To U
    B = Split(A(J), "|") ' Splitting B(0)=Pj | B(1)=QDte
    M.Pj = B(0)
    M.QDte = B(1)
    O(J) = M    '<===
Next
ZPjKey = O
End Property

Private Function TKeyDta(Pj$, Sku$, QDte As Date) As KeyDta
Dim O As KeyDta
O.Pj = Pj
O.QDte = QDte
O.Sku = Sku
TKeyDta = O
End Function

Private Function TPjKey(Pj$, QDte As Date) As PjKey
Dim O As PjKey
O.Pj = Pj
O.QDte = QDte
TPjKey = O
End Function

Private Sub ZAssert_NewQDte_ShouldBeNewest(NewQDte As Date)
Dim MaxQDte As Date
    MaxQDte = RunSql_Val(Db, "Select Max(QuoteDate) from ProjQ")
    If MaxQDte > NewQDte Then Er "The given {NewQuoteDate} is smaller than {LatestQuoteDate} in database", NewQDte, MaxQDte
End Sub

Private Sub ZAssert_PjKeyCannotDup(PjKey() As PjKey)
Dim M1 As PjKey, A$(), B$, J&
ReDim A(UBound(PjKey))
For J = 0 To UBound(PjKey)
    With PjKey(J)
        B = .Pj & "|" & .QDte
    End With
    If Ay_Has(A, B) Then Stop
    A(J) = B
Next
End Sub

Private Sub ZCpy_TblOne(PjWhere$)
Z_Run "Insert into ProjOneTimeCost" & _
" Select ProjNo,#{0}# as QuoteDate, OneTimeCost," & _
" Cost, OneTimeCostRmk" & _
" From ProjOneTimeCost" & _
" where {1}", NewQDte, PjWhere '<== Cpy from 1 record from Old-Key to New-Key
End Sub

Private Sub ZCpy_TblPjQ(PjWhere$)
Dim MassUpdDtaFxFInfo As FInfo   ' MassUpdDtaFInfo
    MassUpdDtaFxFInfo = Ffn_FInfo(Src_Fx)
    
Dim Fn$, Sz$, Tim As Date
With MassUpdDtaFxFInfo
    Fn = .Ffn
    Sz = .Sz
    Tim = .Tim
End With
'-- Cpy
Z_Run "Insert into ProjQ" & _
" Select ProjNo,#{0}# as QuoteDate, Supplier," & _
" '{1}' as QpsFormFileName, {2} as QpsFormFileLen, #{3}# as QpsFormFileDateTime," & _
" RateUSD, RateCHF, RateJPY" & _
" From ProjQ" & _
" where {4}", NewQDte, Fn, Sz, Tim, PjWhere
End Sub

Private Sub ZCpy_TblSku(SkuWhere$)
'-- Cpy
Z_Run "Insert into Sku Select x.ProjNo,#{0}# as QuoteDate, x.Sku, x.PotentialQty, x.Cost," & _
" IIf(IsNull(x.CompleteWatchUSD),0,x.CompleteWatchUSD) as CompleteWatchUSD," & _
" IIf(IsNull(x.CompleteWatchHKD),0,x.CompleteWatchHKD) as CompleteWatchHKD," & _
" IIf(IsNull(x.AssWatchUSD)     ,0,x.AssWatchUSD)      as AssWatchUSD     ," & _
" IIf(IsNull(x.AssWatchHKD)     ,0,x.AssWatchHKD)      as AssWatchHKD     ," & _
" IIf(IsNull(x.SalesmanUSD)     ,0,x.SalesmanUSD)      as SalesmanUSD     ," & _
" IIf(IsNull(x.SalesmanHKD)     ,0,x.SalesmanHKD)      as SalesmanHKD" & _
" From Sku x" & _
" Where {1}", NewQDte, SkuWhere
End Sub

Private Sub ZCpy_TblSkuCostChr(SkuWhere$)
Z_Run "Insert into SkuCostChr" & _
" Select ProjNo,#{0}# as QuoteDate, Sku," & _
" CostGp, CostEle, CharCode, CharVal" & _
" From SkuCostChr" & _
" Where {1}", NewQDte, SkuWhere
End Sub

Private Sub ZCpy_TblSkuCostEle(SkuWhere$)
Z_Run "Insert into SkuCostEle" & _
" Select ProjNo,#{0}# as QuoteDate, Sku," & _
" CostGp, CostEle, Cost, CostEleRmk" & _
" From SkuCostEle" & _
" where {1}", NewQDte, SkuWhere
End Sub

Private Function ZPjWhere$(P As PjKey)
ZPjWhere = Fmt_QQ("ProjNo='?' and QuoteDate=#?#", P.Pj, P.QDte)
End Function

Private Sub ZTst_Assert_RecCnt(ToWherePj$, ToWhereSku$, ExpRecCntAy&())
Dim J%
Dim Msg$
Dim O$()
Dim ActRecCntAy&()
Dim TblNmAy$()
For J = 0 To UB(ExpRecCntAy)
    If ActRecCntAy(J) <> ExpRecCntAy(J) Then
        Msg = Fmt_QQ("RecCnt of Tbl(?) ExpRecCnt(?) <> ActRecCnt(?)", TblNmAy(J), ActRecCntAy(J), ExpRecCntAy(J))
        Push O, Msg
    End If
Next
Assert_AyIsEmpty O
End Sub

Private Sub ZTst_FndTstDta(NewQDte As Date, ToWherePj$, ToWhereSku$, KeyDta() As KeyDta, RecCntAy_Fm&())
Dim Pj$
Dim SkuAy$(2)
Dim J%
Dim PjWhere$
Dim QDte As Date

'-- A Pj with 6 SKU ------------------------
With Db.OpenRecordset("Select ProjNo, QuoteDate, Count(*) as Cnt from Sku group by ProjNo, QuoteDate Having Count(*)=6")
    If Not .EOF Then
        Pj = !ProjNo
        QDte = !QuoteDate
    End If
    .Close
End With
Stop
'FmPj = TPjKey(Pj, QDte)     ' <===

'-- Take 3 Sku -----------------------------------
PjWhere = Fmt_QQ("ProjNo='?' and QuoteDate=#?#", Pj, QDte)

With Db.OpenRecordset(Fmt_QQ("Select Sku from Sku where ?", PjWhere))
    SkuAy(0) = !Sku '<===
    .MoveNext
    SkuAy(1) = !Sku '<===
    .MoveNext
    SkuAy(2) = !Sku '<===
    .Close
End With

'-- Fnd NewQDte ----------------------------------
NewQDte = RunSql_Val(Db, "Select Max(QuoteDate) from ProjQ")
NewQDte = DateAdd("D", 1, NewQDte)
'-- Set KeyDta 3 elements
ReDim KeyDta(2)
For J = 0 To 2
    With KeyDta(J)  '<===
        .Pj = Pj
        .QDte = NewQDte
        .Sku = SkuAy(J)
    End With
Next
'-- Fnd ToPj -------------------------------
Stop
'ToPj = TPjKey(Pj, NewQDte) '<====

'-- Fnd RecCntAy
'
Stop
'RecCntAy = ZTst_RecCntAy(TblNmLvs, PjWhere, SkuWhere)
End Sub

Private Function ZTst_RecCntAy(PjWhere$, SkuWhere$) As Long()
Dim A&()
Dim B&()
A = ZTst_RecCntAyOneWhere(TblNmLvs_Pj_, PjWhere)
B = ZTst_RecCntAyOneWhere(TblNmLvs_Sku, SkuWhere)
ZTst_RecCntAy = Ay_Add(A, B)
End Function

Private Function ZTst_RecCntAyOneWhere(TblNmLvs$, Where$) As Long()
Dim J%
Dim Ay$()
Dim O&()
Dim U&

Ay = SplitLvs(TblNmLvs)
U = UB(Ay)
ReDim O(U)
For J = 0 To U
    O(J) = ZTst_RecCntTbl(Ay(J), Where)
Next
ZTst_RecCntAyOneWhere = O
End Function

Private Function ZTst_RecCntTbl&(TblNm$, Where$)
Dim Sql$
Sql = Fmt_QQ("Select Count(*) from ? where ?", TblNm, Where)
ZTst_RecCntTbl = RunSql_Val(Db, Sql)
End Function

Private Sub Z_Run(S$, ParamArray Ap())
Dim Av()
    Av = Ap
Dim Sql$
    Sql = Fmt_Av(S, Av)
Db.Execute Sql
End Sub

Private Sub Z_RunDlt(Where$, TblNmLvs$)
Dim SqlAy$()
Dim Sql$
Dim J%
Dim TblNmAy$()
    TblNmAy = SplitLvs(TblNmLvs)
For J = 0 To UB(TblNmAy)
    Z_Run "Delete from {0} where {1}", TblNmAy(J), Where
Next
End Sub

