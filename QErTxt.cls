VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QErTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function ChrCdNotFnd(D As MsgDta_ChrCdNotFnd) As TErTxt
ChrCdNotFnd.Sht = "Invalid Char Code"
ChrCdNotFnd.Msg = Fmt_QQ("A char column header has a invalid value[?]", D.ErChrCd)
End Function

Function ChrEmpty() As TErTxt
ChrEmpty.Sht = "Empty Char"
ChrEmpty.Msg = Fmt_QQ("The charactertistic must be entered.")
End Function

Function ChrVal() As TErTxt
ChrVal.Sht = "Invalid Char Val"
ChrVal.Msg = Fmt_QQ("The characteristic value entered is not in valid list of Characteristic.")
End Function

Function DifColCnt(D As MsgDta_DifColCnt) As TErTxt
Const C = "Worksheet(Original).and.(Working) has different #of-Columns({0}).and.({1})."
DifColCnt.Msg = Fmt(C, D.OrgHdColSz, D.WrkHdColSz)
DifColCnt.Sht = "Col Count Dif"
End Function

Function DifHdCell(D As MsgDta_DifHdCell) As TErTxt
Const C = "Header cell of worksheet(Original).and.(Working) have different value({0}).and.({1})."
DifHdCell.Msg = Fmt(C, D.OrgHdVal, D.WrkHdVal)
DifHdCell.Sht = "Head Cell Dif"
End Function

Function DifR1Formula(D As MsgDta_DifR1Formula) As TErTxt
DifR1Formula.Sht = "Dif from Row-1"
DifR1Formula.Msg = Fmt_QQ("This cell has formula[?] which must equal to Row-1-formula[?]", D.ErFormula, D.R1Formula)
End Function

Function DifVal(D As MsgDta_DifVal) As TErTxt
DifVal.Sht = "Dif from [Original]"
DifVal.Msg = Fmt_QQ("The [?] field should not be changed.  It must equal to [?] in [Original] at [?]", D.FldNm, D.OrgVal, D.OrgAdr)
End Function

Function DupSku(D As MsgDta_DupSku) As TErTxt
DupSku.Sht = "Duplicate SKU"
DupSku.Msg = Fmt_QQ("This Sku is duplicate at row-[?]", D.FirstRno_WithDupSku& + 6)
End Function

Function NoOrgRow() As TErTxt
NoOrgRow.Sht = "Not in [Original]"
NoOrgRow.Msg = "This Sku [Pj+QuoteDate+Sku] does not found in [Original] worksheet"
End Function

Function ValTy(D As MsgDta_ValTy) As TErTxt
ValTy.Sht = "Value type error"
ValTy.Msg = Fmt_QQ("This cell has unexpected value type(?) of value[?].  It should have this value type(?)", TypeName(D.ErVal), D.ErVal, D.ExpDtaTy)
End Function
