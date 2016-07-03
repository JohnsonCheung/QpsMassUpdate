Attribute VB_Name = "DoImport"
Option Explicit
Private Db As Database
Private NewQDte As Date
Private MassUpd_Wb As Workbook
Private DtaChg() As TDtaChg

Sub Do_Import()
Dim Fx$
Set MassUpd_Wb = Src_Wb
Fct.AlignTwoWb MassUpd_Wb
ZAssert_NoEr
DtaChg = Src_DtaChg
ZAssert_AnyDtaChg
Set Db = Dao.DBEngine.OpenDatabase(Fs_Fb)
ZNewQDte_Set
If NewQDte = 0 Then Db.Close: Exit Sub
ZDoCpy_ToNewQuote
Dim R&
For R = 0 To UBound(DtaChg)
    ZDoUpd__OneDtaChg DtaChg(R)
Next
ZDoDlt_ZerCstRec_In_ProjOneTimeCost_and_SkuCostEle
Db.Close
Fx = MassUpd_Wb.FullName
MassUpd_Wb.Close False
ZDoMov_MassUpdFx Fx
ZDoShellFb_ToGenCmpRpt
End Sub

Private Sub ZAssert_AnyDtaChg()
If DtaChg_Sz(DtaChg) Then Exit Sub
Er "No data change"
End Sub

Private Sub ZAssert_NoEr()
Dim A As TDtaErOpt
    A = Src_DtaEr

If A.Some > 0 Then
    Er "There are {N} validation error.  Cannot import.  Use valid to show all the error and make correction before import", UBound(A.Ay) + 1
End If
End Sub


Private Sub ZDoCpy_ToNewQuote()
'=== Copy Pj & Sku in DtaChg to NewQuote ==========================

'Cpy array of quote from Old-Quote-Date to NewQDte
'Inside PjKey array, it has all the Pj+QDte to copy to NewQDte.
'                    2-tables use this Key: PjQ & One
'Inside KeyDta array, it has all the Pj+QDte+Sku to copy to NewQDte.
'                     3-tables use this Key: Sku & SkUCostEle & SkuCostChr

Dim KeyDta() As KeyDta
    KeyDta = DtaChg_KeyDta(DtaChg)

Do_CpyQuote Db, KeyDta, NewQDte
End Sub

Private Sub ZDoDlt_ZerCstRec_In_ProjOneTimeCost_and_SkuCostEle()
Db.Execute "Delete from ProjOneTimeCost where Cost=0" ' Delete any Cost=0 after the update
Db.Execute "Delete from SkuCostEle where Cost=0"
End Sub

Private Sub ZDoMov_MassUpdFx(FmFx$)
Dim ToFdr$
Dim Fn$
Dim Fdr$
Dim SubFdr$

SubFdr = "NewQuote " & Format(NewQDte, "yyyy-mm-dd") & "\"
ToFdr = Fs_ImpDoneFdr & SubFdr
Pth_CrtIfNotExist ToFdr
Fso.MoveFile FmFx, ToFdr    '<===
Fn = Ffn_Fn(FmFx)
Fdr = Ffn_Pth(FmFx)
Msg "{Mass-Update-Xls-File} in {Fdr} is imported into database and moved to {Sub-Fdr}", Fn, Fdr, SubFdr
End Sub

Private Sub ZDoShellFb_ToGenCmpRpt()
Dim WrkFdr$
Dim NewQQryFxFn$
Dim MassUpdFxFn$
Stop
Setting.GenCmpRpt_Sav NewQDte, WrkFdr, NewQQryFxFn, MassUpdFxFn
Shell Fs_Fb
End Sub

Private Sub ZDoUpd(FmtStr$, ParamArray Ap())
Dim Sql$
Dim Av()
Av = Ap
Sql = Fmt_Av(FmtStr, Av)
Db.Execute Sql
End Sub

Private Sub ZDoUpd_ChgOfChr(P As TDtaChg)
'== Delete rec ================================
Dim IsWrkVal_Blank As Boolean
    IsWrkVal_Blank = Trim(P.WrkVal) = ""

Dim Where$
    Const W$ = "ProjNo='?' and QuoteDate=#?# and Sku='?' and CostGp='?' and CostEle='?' and CharCode='?'"
    With P
        Where = Fmt_QQ(W, .Key.Pj, NewQDte, .Key.Sku, .CostGp, .CostEle, .CharCode)
    End With
    
Dim Sql_Dlt$
    If IsWrkVal_Blank Then
        Sql_Dlt = "Delete from SkuCostChr where " & Where
    End If

If IsWrkVal_Blank Then
    Db.Execute Sql_Dlt '<===
    Exit Sub
End If

'== SqlIsRec_Exist record in the ProjQ ==============
Dim IsRec_Exist As Boolean
    Dim A$
    
    A = "Select Count(*) from SkuCostChr where " & Where
    IsRec_Exist = RunSql_Val(Db, A) > 0

Dim Sql_Ins_or_Upd$
    With P
        If IsRec_Exist Then
            Sql_Ins_or_Upd = Fmt_QQ("Update SkuCostChr set CharVal='?' where ?", .WrkVal, Where)
        Else
            Const SqlIns = "Insert into SkuCostChr (ProjNo,QuoteDate,Sku,CostGp,CostEle,CharCode,CharVal) values ('?',#?#,'?','?','?','?','?')"
            Sql_Ins_or_Upd = Fmt_QQ(SqlIns, .Key.Pj, NewQDte, .Key.Sku, .CostGp, .CostEle, .CharCode, .WrkVal)
        End If
    End With
ZDoUpd Sql_Ins_or_Upd    '<===
End Sub

Private Sub ZDoUpd_ChgOfCstRmk(P As TDtaChg)
Dim Where$
    Where = Fmt("ProjNo='{0}' and QuoteDate=#{1}# and Sku='{2}' and CostGp='{3}' and CostEle='{4}'", P.Key.Pj, NewQDte, P.Key.Sku, P.CostGp, P.CostEle)

Dim IsRec_Exist As Boolean
    Dim A$
    A = Fmt_QQ("Select Count(*) from SkuCostEle where ?", Where)
    IsRec_Exist = RunSql_Val(Db, A) > 0

Dim Sql$
    With P
        If IsRec_Exist Then
            Sql = Fmt_QQ("Update SkuCostEle set CostEleRmk='?' where ?", .WrkVal, Where)
        Else
            Const SqlIns = "Insert into SkuCostEle (ProjNo,QuoteDate,Sku,CostGp,CostEle,CostEleRmk) values ('?',#?#,'?','?','?','?')"
            Sql = Fmt_QQ(SqlIns, .Key.Pj, NewQDte, .Key.Sku, .CostGp, .CostEle, .WrkVal)
        End If
    End With
ZDoUpd Sql
End Sub

Private Sub ZDoUpd_ChgOfCstVal(P As TDtaChg)
Dim Where$
    Where = Fmt("ProjNo='{0}' and QuoteDate=#{1}# and Sku='{2}' and CostGp='{3}' and CostEle='{4}'", P.Key.Pj, NewQDte, P.Key.Sku, P.CostGp, P.CostEle)
Dim IsRec_Exist As Boolean
    Dim A$
    A = Fmt("Select Count(*) from SkuCostEle where ?", Where)
    IsRec_Exist = RunSql_Val(Db, A) > 0

Dim Sql_Ins_or_Upd$
    With P
        If IsRec_Exist Then
            Sql_Ins_or_Upd = Fmt_QQ("Update SkuCostEle set Cost='?' where ?", .WrkVal, Where)
        Else
            Dim Fld$
                Select Case P.FldNm
                Case "Cost": Fld = "Cost"
                Case "CostEleRmk": Fld = "CostEleRmk"
                Case Else: Stop
                End Select
            Const SqlIns = "Insert into SkuCostEle (ProjNo,QuoteDate,Sku,CostGp,CostEle,?) values ('?',#?#,'?','?','?','?')"
            Sql_Ins_or_Upd = Fmt_QQ(SqlIns, Fld, .Key.Pj, NewQDte, .Key.Sku, .CostGp, .CostEle, .WrkVal)
        End If
    End With
ZDoUpd Sql_Ins_or_Upd
End Sub

Private Sub ZDoUpd_ChgOfOne(P As TDtaChg)
Dim Fld$, V$, OneTimeCost$, Pj$
    Select Case P.FldNm
    Case "ProtRmk": Fld = "OneTimeCostRmk": OneTimeCost = "Prototype Cost"
    Case "ProtCst": Fld = "Cost"::::::::::: OneTimeCost = "Prototype Cost"
    Case "ToolRmk": Fld = "OneTimeCostRmk": OneTimeCost = "Tooling Cost"
    Case "ToolCst": Fld = "Cost"::::::::::: OneTimeCost = "Tooling Cost"
    Case Else: Stop
    End Select
    Pj = P.Key.Pj
    V = P.WrkVal

Dim IsRec_Exist As Boolean
    Dim A$
    A = Fmt("Select Count(*) from ProjOneTimeCost where ProjNo='{0}' and QuoteDate=#{1}# and OneTimeCost='{2}'", Pj, NewQDte, OneTimeCost)
    IsRec_Exist = RunSql_Val(Db, A) > 0

Dim Sql_Ins_or_Upd$
    If IsRec_Exist Then
        Sql_Ins_or_Upd = Fmt("Update ProjOneTimeCost set {0}='{1}' where ProjNo='{2}' and QuoteDate=#{3}# and OneTimeCost='{4}'", _
            Fld, V, Pj, NewQDte, OneTimeCost)
    Else
        Sql_Ins_or_Upd = Fmt("insert into ProjOneTimeCost (ProjNo, QuoteDate, OneTimeCost, {0}) values ('{1}',#{2}#,'{3}','{4}')", _
            Fld, Pj, NewQDte, OneTimeCost, V)
    End If

ZDoUpd Sql_Ins_or_Upd
End Sub

Private Sub ZDoUpd_ChgOfPjQ(P As TDtaChg)
Dim Sql$
    Dim Fld$, V$, Pj$
        Pj = P.Key.Pj
        Select Case P.FldNm
        Case "RateCHF"
        Case "RateUSD"
        Case "RateJPY"
        Case Else: Stop
        End Select
    Fld = P.FldNm
    V = P.WrkVal
ZDoUpd "Update ProjQ set {0}='{1}' where ProjNo='{2}' and QuoteDate=#{3}#", Fld, V, Pj, NewQDte
End Sub

Private Sub ZDoUpd_ChgOfSku(P As TDtaChg)
'The new-quote-sku-rec should exist, due to it is an "Update" or a "Copy", which is done in Cpy_Quote
Dim FldNm$, V$, Pj$, Sku$
    Select Case P.FldNm
    Case "AssWatchHKD"
    Case "AssWatchUSD"
    Case "CompleteWatchHKD"
    Case "CompleteWatchUSD"
    Case "Cost"
    Case "PotentialQty"
    Case "SalesmanHKD"
    Case "SalesmanUSD"
    Case Else: Stop
    End Select
    FldNm = P.FldNm
    V = P.WrkVal
    Pj = P.Key.Pj
    Sku = P.Key.Sku

    '== Update the field =========================
ZDoUpd "Update Sku set {0}='{1}' where ProjNo='{2}' and QuoteDate=#{3}# and Sku='{4}'", FldNm, V, Pj, NewQDte, Sku
End Sub

Private Sub ZDoUpd__OneDtaChg(M As TDtaChg)
Select Case M.FldTy
Case eFldTy.ePjQ:    ZDoUpd_ChgOfPjQ M            ' Always update, no need to handle add or delete.  The delete is done by 2 delete sql below
Case eFldTy.eOne:    ZDoUpd_ChgOfOne M
Case eFldTy.eSku:    ZDoUpd_ChgOfSku M
Case eFldTy.eCstVal: ZDoUpd_ChgOfCstVal M
Case eFldTy.eCstRmk: ZDoUpd_ChgOfCstRmk M
Case eFldTy.eChr:    ZDoUpd_ChgOfChr M
Case Else: Stop
End Select
End Sub

Private Sub ZNewQDte_Set()
NewQDte = ZNewQDte_Ask
If NewQDte = 0 Then MsgBox "No import", vbInformation
End Sub

Private Function ZNewQDte_Ask() As Date
Const Fmt = "yyyy-mm-dd"

Dim MaxQDte As Date
Dim MinQDte As Date
    MaxQDte = RunSql_Val(Db, "Select Max(QuoteDate) from ProjQ")
    MinQDte = DateAdd("D", 1, Max(MaxQDte, Date))

Dim A$
    On Error Resume Next
Again:
    A = InputBox("Input a new quote date" & vbLf & ">=" & Format(MinQDte, Fmt), , Format(MinQDte, Fmt))
    If A = "" Then Exit Function
    If Not IsDte(A) Then MsgBox Fmt_ErDes("Invalid {date}", Array(A)), vbCritical: GoTo Again
    If A < MinQDte Then
        MsgBox "Date must >= " & Format(MinQDte, Fmt), vbCritical
        GoTo Again
    End If
    
ZNewQDte_Ask = A
Exit Function
R1:
    MsgBox "Invalid date", vbCritical: GoTo Again
End Function


