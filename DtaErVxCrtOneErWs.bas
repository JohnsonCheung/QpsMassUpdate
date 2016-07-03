Attribute VB_Name = "DtaErVxCrtOneErWs"
Option Explicit
Type TErMsg1
    Key As KeyDta
    FldNm As String
    CostGp As String
    CostEle As String
    CharName As String
    Ws As String
    Adr As String     ' Adr in Ws with error.  If Ws="", Working is assumed
'   A_LstAdr As String  'No LstAdr in single error worksheet
'   A_LnkAdr As String
    ErVal As String     ' The error value entered
    OrgVal As String    ' The value from OrgWs, for reference only
    ErTxt As TErTxt
    Must As String
    Multi As String
End Type

Private A_DtaEr As TDtaErOpt
Private A_ErVer As eErWsVer
Private Type B
    ErWs As Worksheet
    ErMsg() As TErMsg1
End Type
Private B As B

Sub DtaErVx_Crt_OneErWs(Wb As Workbook, DtaEr As TDtaErOpt, ErVer As eErWsVer)
If IsNothing(Wb) Then Exit Sub
A_DtaEr = DtaEr
A_ErVer = ErVer

Select Case ErVer
Case eV2DropDown, eV3SelInEr, eV4SelInSep
Case Else: Er "Given {ErVer} should be V2..V4", ErVer
End Select

Wb_DltWs Wb, ErWsNm
Wb_DltWs Wb, LstWsNm
Wb_DltWs Wb, "Selection"
If Not DtaEr.Some Then Exit Sub

Application.ScreenUpdating = False

Set B.ErWs = Wb_AddWs_AtEnd(Wb, ErWsNm)   '<== Crt ErWs
B.ErMsg = ZErMsg

ZErWs_DoPut_ErMsg
ZErWs_DoLnk_ColWrkAdr_ToWrkWs
ZErWs_DoCrt_Macro
ZSrcWs_DoLnk_ErCell_ToErWs
Application.ScreenUpdating = True
End Sub

Sub ZErWs_DoPut_ErMsg()
Dim ErMsg() As TErMsg1
Dim Ws As Worksheet
    Set Ws = B.ErWs
    ErMsg = B.ErMsg
Dim R2&
    R2 = UBound(ErMsg) + 1
Ws_CRR(Ws, ZErCno_QDte, 2, R2).NumberFormat = "yyyy-mm-dd"
Ws_CRR(Ws, ZErCno_ErVal, 2, R2).NumberFormat = "@"
Ws_CRR(Ws, ZErCno_Sku, 2, R2).NumberFormat = "@"

With Ws
    Cell_PutSqv .Range("A1"), ZErWs_HdSqv
    Cell_PutSqv .Range("A2"), ZErWs_Sqv
End With

Dim A As Range
    Set A = Ws.Cells(2, ZErCno_Must)

Cell_Freeze A

ZErWs_DoSet_OutLIne_Lvl2 Ws, ZErCno_Msg
ZErWs_DoSet_OutLIne_Lvl2 Ws, ZErCno_CostGp
ZErWs_DoSet_OutLIne_Lvl2 Ws, ZErCno_CostEle

Select Case A_ErVer
Case eErWsVer.eV2DropDown:  ZErWs_DoSet_DropDown_SelCol
Case eErWsVer.eV3SelInEr
Case eErWsVer.eV4SelInSep
Case Else:                  Er "Given {ErVer} is invalid", A_ErVer
End Select

Ws_R(Ws, 1).AutoFilter
Ws.Columns.AutoFit
Ws_Zoom Ws, 85
Ws_RR(Ws, 1, Ws_LastRow(Ws)).VerticalAlignment = xlVAlignCenter
Ws.Outline.SummaryColumn = xlSummaryOnLeft
'Ws.Protect AllowFiltering:=True, AllowFormattingColumns:=True

End Sub

Private Function ZDtaEr() As TDtaEr()
ZDtaEr = A_DtaEr.Ay
End Function

Private Property Get ZDtaEr_Itm(Idx&) As TDtaEr
ZDtaEr_Itm = A_DtaEr.Ay(Idx)
End Property

Private Property Get ZDtaEr_Sz&()
If A_DtaEr.Some Then
    ZDtaEr_Sz = UBound(A_DtaEr.Ay) + 1
End If
End Property

Private Function ZErCno%(FldNm$)
Dim O%
O = Ay_Idx(ZErCno_FldNmAy, FldNm) + 1
If O = 0 Then Stop
ZErCno = O
End Function

Private Property Get ZErCno_CharName%()
ZErCno_CharName = ZErCno("CharName")
End Property

Private Property Get ZErCno_Corr%()
ZErCno_Corr = ZErCno("Correction")
End Property

Private Property Get ZErCno_CostEle%()
ZErCno_CostEle = ZErCno("CostEle")
End Property

Private Property Get ZErCno_CostGp%()
ZErCno_CostGp = ZErCno("CostGp")
End Property

Private Property Get ZErCno_ErVal%()
ZErCno_ErVal = ZErCno("ErVal")
End Property

Private Property Get ZErCno_FldLst$()
Const A_FldLst1 = "Sht Msg Pj Sku QDte FldNm CostGp CostEle CharName Must Multi Ws WrkAdr ErVal Correction Selection"
Const A_FldLst2 = "Sht Msg Pj Sku QDte FldNm CostGp CostEle CharName Must Multi Ws WrkAdr ErVal Correction"
Dim O$
Select Case A_ErVer
Case eV1ErAndLst:               Er "Program Logic Error"
Case eV2DropDown, eV3SelInEr:   O = A_FldLst1
Case eV4SelInSep:               O = A_FldLst2
Case Else: Er "Pgm error"
End Select
ZErCno_FldLst = O
End Property

Private Property Get ZErCno_FldNm%()
ZErCno_FldNm = ZErCno("FldNm")
End Property

Private Function ZErCno_FldNmAy() As String()
ZErCno_FldNmAy = Split(ZErCno_FldLst)
End Function

Private Property Get ZErCno_Msg%()
ZErCno_Msg = ZErCno("Msg")
End Property

Private Property Get ZErCno_Must%()
ZErCno_Must = ZErCno("Must")
End Property

Private Property Get ZErCno_QDte%()
ZErCno_QDte = ZErCno("QDte")
End Property

Private Property Get ZErCno_Sel%()
ZErCno_Sel = ZErCno("Selection")
End Property

Private Property Get ZErCno_Sku%()
ZErCno_Sku = ZErCno("Sku")
End Property

Private Property Get ZErCno_WrkAdr%()
ZErCno_WrkAdr = ZErCno("WrkAdr")
End Property

Private Property Get ZErCno_Ws%()
ZErCno_Ws = ZErCno("Ws")
End Property

Private Function ZErMsg() As TErMsg1()
Dim NEr&
    NEr = ZDtaEr_Sz

If NEr = 0 Then Exit Function

Dim DtaEr() As TDtaEr
    DtaEr = ZDtaEr

Dim O() As TErMsg1
    ReDim O(NEr - 1)
    Dim I&
    For I = 0 To NEr - 1
        O(I) = ZErMsg_One(DtaEr(I))
    Next

ZErMsg = O
End Function

Private Function ZErMsg_ChrCdNotFnd(Er As DE_ChrCdNotFnd) As TErMsg1
Dim O As TErMsg1
O.ErTxt = QErTxt.ChrCdNotFnd(Er.MsgDta)
With Er.ShwFld
    O.Adr = .Adr
    O.CharName = .CharName
'    O.A_LnkAdr = LnkAdr
    O.CostEle = .CostEle
    O.CostGp = .CostGp
    O.FldNm = .FldNm
End With
ZErMsg_ChrCdNotFnd = O
End Function

Private Function ZErMsg_ChrEmpty(Er As DE_ChrEmpty) As TErMsg1
Dim O As TErMsg1
O.ErTxt = QErTxt.ChrEmpty()
With Er.ShwFld
    O.Ws = .Ws
    O.Adr = .Adr
    O.CharName = .CharName
    O.CostEle = .CostEle
    O.CostGp = .CostGp
    O.FldNm = "Char" ' FldNm_OfChr(.CostGp, .CostEle, .CharName)
    O.Key = .Key
    O.OrgVal = .OrgVal
End With
ZErMsg_ChrEmpty = O
End Function

Private Function ZErMsg_ChrVal(Er As DE_ChrVal) As TErMsg1
Dim O As TErMsg1
O.ErTxt = QErTxt.ChrVal()
With Er.ShwFld
    O.Adr = .Adr
    O.CharName = .CharName
    O.CostEle = .CostEle
    O.CostGp = .CostGp
    O.FldNm = "Char" ' FldNm_OfChr(.CostGp, .CostEle, .CharName)
    O.Key = .Key
    O.ErVal = .ErVal
    O.OrgVal = .OrgVal
    '.Ws
End With
ZErMsg_ChrVal = O
End Function

Private Function ZErMsg_DifColCnt(Er As DE_DifColCnt) As TErMsg1
Dim O As TErMsg1
O.ErTxt = QErTxt.DifColCnt(Er.MsgDta)
With Er.ShwFld
     O.Adr = .Adr
     O.Ws = .WsNmWhichIsLargerNoOfCol
'    O.ChrLnkAdr = ChrLnkAdr
'    O.CostEle = .CostEle
'    O.CostGp = .CostGp
'    O.FldNm = FldNm_OfChr(.CostGp, .CostEle, .CharName)
'    O.Pj = .Key.Pj
'    O.QDte = .Key.QDte
'    O.Sku = .Key.Sku
    '.Ws
End With
ZErMsg_DifColCnt = O
End Function

Private Function ZErMsg_DifHdCell(Er As DE_DifHdCell) As TErMsg1
Dim O As TErMsg1
O.ErTxt = QErTxt.DifHdCell(Er.MsgDta)
With Er.ShwFld
    O.Adr = .Adr

'    O.CharName = .CharName
'    O.CostEle = .CostEle
'    O.CostGp = .CostGp
'    O.FldNm = FldNm_OfChr(.CostGp, .CostEle, .CharName)
'    O.Pj = .Key.Pj
'    O.QDte = .Key.QDte
'    O.Sku = .Key.Sku
    'O.Ws
End With
ZErMsg_DifHdCell = O
End Function

Private Function ZErMsg_DifR1Formula(Er As DE_DifR1Formula) As TErMsg1
Dim O As TErMsg1
O.ErTxt = QErTxt.DifR1Formula(Er.MsgDta)
With Er.ShwFld
    O.Adr = .Adr
    O.CostEle = .CostEle
    O.CostGp = .CostGp
    O.FldNm = .FldNm
    O.Key = .Key
    O.Ws = .Ws
End With
ZErMsg_DifR1Formula = O
End Function

Private Function ZErMsg_DifVal(Er As DE_DifVal) As TErMsg1
Dim O As TErMsg1
O.ErTxt = QErTxt.DifVal(Er.MsgDta)
With Er.ShwFld
    O.ErVal = .ErVal
    O.Adr = .Adr
    O.FldNm = .FldNm
    O.Key = .Key
End With
ZErMsg_DifVal = O
End Function

Private Function ZErMsg_DupSku(Er As DE_DupSku) As TErMsg1
Dim O As TErMsg1
O.ErTxt = QErTxt.DupSku(Er.MsgDta)
With Er.ShwFld
    O.Key = .Key
    O.Adr = .Adr
    O.FldNm = .FldNm
End With
ZErMsg_DupSku = O
End Function

Private Function ZErMsg_NoOrgRow(Er As DE_NoOrgRow) As TErMsg1
Dim O As TErMsg1
O.ErTxt = QErTxt.NoOrgRow
With Er.ShwFld
    O.Key = .Key
    O.FldNm = .FldNm
    O.Adr = .Adr
End With
ZErMsg_NoOrgRow = O
End Function

Private Function ZErMsg_One(Er As TDtaEr) As TErMsg1
Dim O As TErMsg1, T As TErTxt
Select Case Er.Ty
Case eDtaErTy.eChrValEr:       O = ZErMsg_ChrVal(Er.ChrVal)
    If Er.ChrVal.ChrDef.IsMust Then O.Must = "Must"
    If Er.ChrVal.ChrDef.IsMulti Then O.Multi = "Multi"
Case eDtaErTy.eChrEmptyEr:     O = ZErMsg_ChrEmpty(Er.ChrEmpty)
    If Er.ChrEmpty.ChrDef.IsMust Then O.Must = "Must"
    If Er.ChrEmpty.ChrDef.IsMulti Then O.Multi = "Multi"
Case eDtaErTy.eChrCdNotFndEr:  O = ZErMsg_ChrCdNotFnd(Er.ChrCdNotFnd)
Case eDtaErTy.eDifHdCellEr:    O = ZErMsg_DifHdCell(Er.DifHdCell)
Case eDtaErTy.eDifColCntEr:    O = ZErMsg_DifColCnt(Er.DifColCnt)
Case eDtaErTy.eDifR1FormulaEr: O = ZErMsg_DifR1Formula(Er.DifR1Formula)
Case eDtaErTy.eDifValEr:       O = ZErMsg_DifVal(Er.DifVal)
Case eDtaErTy.eDupSkuEr:       O = ZErMsg_DupSku(Er.DupSku)
Case eDtaErTy.eNoOrgRowEr:     O = ZErMsg_NoOrgRow(Er.NoOrgRow)
Case eDtaErTy.eValTyEr:        O = ZErMsg_ValTy(Er.ValTy)
Case Else: Stop
End Select
ZErMsg_One = O
End Function

Private Function ZErMsg_ValTy(Er As DE_ValTy) As TErMsg1
Dim O As TErMsg1
O.ErTxt = QErTxt.ValTy(Er.MsgDta)
With Er.ShwFld
    O.Adr = .Adr
    O.Key = .Key
    O.FldNm = .FldNm
    O.CostGp = .CostGp
    O.CostEle = .CostEle
    O.CharName = .CharName
    O.ErVal = CStr(.ErVal)
End With
ZErMsg_ValTy = O
End Function

Private Sub ZErWs_DoCrt_Macro()
Dim Fn$, Fn1$
    Select Case A_ErVer
    Case eV2DropDown: Fn = "ErWs_WsChg_V2DropDown"
    Case eV3SelInEr:  Fn = "ErWs_WsChg_V3SelInEr":   Fn1 = "ErWs_WsSelChg_V3SelInEr"
    Case eV4SelInSep: Fn = "ErWs_WsChg_eV4SelInSep": Fn1 = "ErWs_WsSelChg_V4SelInSep"
    Case Else:        Er "Given {ErVer} should be V2..V4", A_ErVer
    End Select
Const Evt1$ = "Change"
Const Evt2$ = "SelectionChange"
Ws_Crt_EvtMth_CallingFn B.ErWs, Evt1, "jjMassUpd", , Fn       '<== Add Worksheet_change
If Fn1 = "" Then Exit Sub
Ws_Crt_EvtMth_CallingFn B.ErWs, Evt2, "jjMassUpd", , Fn1      '<== Optional Add Worksheet_Selectionchange
End Sub

Private Sub ZErWs_DoLnk_ColWrkAdr_ToWrkWs()
Dim ErWs As Worksheet
    Set ErWs = B.ErWs
Dim J%, Rge As Range

Dim WrkWs As Worksheet
Dim OrgWs As Worksheet
    Set WrkWs = Src.Wrk.Ws
    Set OrgWs = Src.Org.Ws

Dim WrkAdrCno%
Dim WsCno%
    WrkAdrCno = ZErCno_WrkAdr
    WsCno = ZErCno_Ws
For J = 2 To Ws_LastRow(ErWs)
    Dim SrcRge As Range
        Set SrcRge = Ws_RC(ErWs, J, WrkAdrCno)
    
    Dim TarAdr$
        TarAdr = SrcRge.Value
    
    If Trim(TarAdr) = "" Then GoTo Nxt
    
    Dim TarWsNm$
        TarWsNm = Ws_RC(ErWs, J, WsCno).Value
        
    Dim TarWs As Worksheet
        Select Case TarWsNm
        Case WrkWsNm, "":   Set TarWs = WrkWs
        Case OrgWsNm:       Set TarWs = OrgWs
        Case Else: Stop
        End Select
    
    Dim TarRge As Range
        Set TarRge = TarWs.Range(TarAdr)
    
    Cell_Lnk SrcRge, TarRge        '<===
Nxt:
Next
End Sub

Private Sub ZErWs_DoSet_DropDown_SelCol()
Dim Cno%
    Cno = ZErCno_Sel
Dim FldNmCno%
    FldNmCno = ZErCno_FldNm
Dim Rno&
For Rno = 2 To ZDtaEr_Sz + 1
    Dim IsNoNeedDropDown As Boolean
        Dim R As Range
            Set R = B.ErWs.Cells(Rno, FldNmCno)
        Dim FldNm$
            FldNm = R.Value
            
        IsNoNeedDropDown = FldNm <> "Char"
    
    If IsNoNeedDropDown Then GoTo Nxt
    
    Dim ListVal$
        ListVal = ZErWs_DropDownListVal(Rno)
    
    If ListVal = "" Then GoTo Nxt
    
    Dim TarCell As Range
        Set TarCell = B.ErWs.Cells(Rno, Cno)
    
    With TarCell.Validation        '<== Set Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=ListVal
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    TarCell.Locked = False
    TarCell.Interior.Color = 10092543
Nxt:
Next
End Sub

Private Sub ZErWs_DoSet_OutLIne_Lvl2(Ws As Worksheet, Cno%)
Ws_C(Ws, Cno).OutlineLevel = 2
End Sub

Private Property Get ZErWs_DropDownListVal$(Rno&)
Dim Idx&
    Idx = Rno - 2
Dim A As TDtaEr
    A = ZDtaEr_Itm(Idx)
    
Dim D As Dictionary
    Select Case A.Ty
    Case eChrValEr:   Set D = A.ChrVal.ChrDef.Dic_OfValNm_ToValCd
    Case eChrEmptyEr: Set D = A.ChrEmpty.ChrDef.Dic_OfValNm_ToValCd
    Case Else: Exit Property
    End Select

Dim B
    B = D.Keys
ZErWs_DropDownListVal = Join(B, ",")
End Property

Private Property Get ZErWs_HdSqv()
ZErWs_HdSqv = Ay_HSqv(ZErCno_FldNmAy)
End Property

Private Property Get ZErWs_Sqv()
Dim Fld$()
    Fld = ZErCno_FldNmAy

Dim NR%
    NR = ZDtaEr_Sz

Dim Msg() As TErMsg1
    Msg = B.ErMsg

Dim NFld%
    NFld = Sz(Fld)
    
ReDim O(1 To NR, 1 To NFld)
    Dim J%, I%, K%
    For J = 0 To NR - 1
        I = 0
        With Msg(J)
            For K = 0 To NFld - 1
                Select Case Fld(K)
'                Case "LnkAdr":    I = I + 1: O(J + 1, I) = .A_LnkAdr
                Case "CharName":  I = I + 1: O(J + 1, I) = .CharName
                Case "WrkAdr":    I = I + 1: O(J + 1, I) = .Adr
                Case "CostEle":   I = I + 1: O(J + 1, I) = .CostEle
                Case "CostGp":    I = I + 1: O(J + 1, I) = .CostGp
                Case "FldNm":     I = I + 1: O(J + 1, I) = .FldNm
                Case "Msg":       I = I + 1: O(J + 1, I) = .ErTxt.Msg
                Case "Pj":        I = I + 1: O(J + 1, I) = .Key.Pj
                Case "QDte":      I = I + 1: O(J + 1, I) = IIf(.Key.QDte = 0, Empty, .Key.QDte)
                Case "Sht":       I = I + 1: O(J + 1, I) = .ErTxt.Sht
                Case "ErVal":     I = I + 1: O(J + 1, I) = .ErVal
                Case "Sku":       I = I + 1: O(J + 1, I) = .Key.Sku
                Case "Ws":        I = I + 1: O(J + 1, I) = .Ws
                Case "Must":      I = I + 1: O(J + 1, I) = .Must
                Case "Multi":     I = I + 1: O(J + 1, I) = .Multi
                Case "Correction", "Selection"
                Case Else: Stop
                End Select
            Next
        End With
    Next
ZErWs_Sqv = O
End Property

Private Sub ZSrcWs_DoLnk_ErCell_ToErWs()
Dim ErMsg() As TErMsg1
Dim ErWs As Worksheet
    ErMsg = B.ErMsg
    Set ErWs = B.ErWs

Dim Wb As Workbook
    Set Wb = B.ErWs.Parent

Dim Org As Worksheet
    Set Org = Wb.Sheets(OrgWsNm)
    
Dim Wrk As Worksheet
    Set Wrk = Wb.Sheets(WrkWsNm)

'Add a lnk to SrcCell which will jmp to ErCell
Wrk.Hyperlinks.Delete   'Assume there is only error link. So clear all should be OK
Org.Hyperlinks.Delete   'Assume
Wrk.ListObjects(1).DataBodyRange.Font.ColorIndex = xlAutomatic
Org.ListObjects(1).DataBodyRange.Font.ColorIndex = xlAutomatic

Dim C%                  ' The AdrCno in ErWrk
    C = ZErCno_WrkAdr

Dim Sz%
    Sz = UBound(ErMsg) + 1

Dim J&
For J = 0 To Sz - 1
    Dim Er As TErMsg1
        Er = ErMsg(J)
    
    If Er.Adr = "" Then GoTo Nxt
    
    Dim WsNm$
        WsNm = Er.Ws
    
    Dim SrcWs As Worksheet
        Select Case WsNm
        Case "", WrkWsNm: Set SrcWs = Wrk
        Case OrgWsNm:     Set SrcWs = Org
        Case Else: Stop
        End Select
    
    Dim SrcCell As Range    ' The cell with err in Wrk or Org ws
        Set SrcCell = SrcWs.Range(Er.Adr)
    
    If Trim(CStr(SrcCell.Value)) = "" Then GoTo Nxt      'If the SrcCell has empty value, cannot create a link
                                            'If create, the cell will have a value of the linked address.
        
    Dim ErCell As Range     ' The Cell in ErWs under the column Adr, which may link to Wrk/Org worksheet
        Set ErCell = ErWs.Cells(J + 2, C)
        'If SrcCell.Address = "$J$320" Then Stop
        
        
    Cell_Lnk SrcCell, ErCell        '<===
    SrcCell.WrapText = True
    SrcCell.VerticalAlignment = xlVAlignTop
    SrcCell.Font.Color = rgbWhite
Nxt:
Next
End Sub
