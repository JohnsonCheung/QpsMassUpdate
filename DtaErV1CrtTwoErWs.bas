Attribute VB_Name = "DtaErV1CrtTwoErWs"
Option Explicit
Const ZLstCno_Fld = "Pj Sku QDte CostGp CostEle CharName IsMust IsMulti WrkAdr ErAdr ErVal CharValName"
Const ZErCno_Fld = "Sht Pj Sku QDte FldNm CostGp CostEle CharName Ws WrkAdr LstAdr ErVal Msg"

Private Type ChrEr   ' Any DtaEr with EmptyChrEr or ChrValEr, it will have one record of this ChrEr, it will used to build LstDr()
    DtaErIdx As Long        ' The Idx of ZDtaEr causing this ChrEr
    Key As KeyDta
    CostGp As String
    CostEle As String
    CharName As String
    CharCode As String
    IsMulti As Boolean
    IsMust As Boolean
    ErVal As String
    WrkAdr As String
    ValDic As Dictionary
End Type
Private Type LstDr
    ChrErIdx As Long    ' The Idx of ChrEr causing this LstDr
    ErAdr As String
    WrkAdr As String
    IsFirst As Boolean
    Key As KeyDta
    CostGp As String
    CostEle As String
    CharName As String
    ErVal As String
    IsMust As Boolean
    IsMulti As Boolean
    CharValName As String
End Type
Private Type TLstInf
    LstAdr() As String
    LstSqv As Variant
End Type

Private Type TErMsg
    Key As KeyDta
    FldNm As String
    CostGp As String
    CostEle As String
    CharName As String
    Ws As String
    Adr As String     ' Adr in Ws with error.  If Ws="", Working is assumed
    A_LstAdr As String
    ErVal As String     ' The error value entered
    OrgVal As String    ' The value from OrgWs, for reference only
    ErTxt As TErTxt
End Type
Private Type B
    ErMsg() As TErMsg
    LstSqv As Variant
    LstWs As Worksheet
    ErWs As Worksheet
End Type
Private A_DtaEr As TDtaErOpt
Private B As B

Sub DtaErV1_DoCrt_TwoErWs(Wb As Workbook, DtaEr As TDtaErOpt)
If IsNothing(Wb) Then Exit Sub
A_DtaEr = DtaEr

Wb_DltWs Wb, ErWsNm
Wb_DltWs Wb, LstWsNm
If Not DtaEr.Some Then Exit Sub
Application.ScreenUpdating = False

Dim A As TLstInf
    A = ZB_ZLstInf
B.LstSqv = A.LstSqv
Set B.LstWs = ZLstWs_DoCrt_EmptyIfNeeded(Wb)
Set B.ErWs = Wb_AddWs_AtEnd(Wb, ErWsNm)
B.ErMsg = ZB_ZErMsg(A.LstAdr)

ZErWs_DoPut_ErMsg
ZErWs_DoLnk_ColWrkAdr_ToSrcWs
ZErWs_DoLnk_ColLstAdr_ToLstWs
ZSrcWs_DoLnk_ErCell_ToErWs
ZLstWs_DoPut_LstSqv
ZLstWs_DoLnk_ColWrk_ToWrkWs
ZLstWs_DoLnk_ColLnk_ToErWs
ZLstWs_DoCrt_ChgSelectionMacro
Application.ScreenUpdating = True
End Sub

Function ZB_ZLstInf() As TLstInf
'Return same # of element as Er.
'Each Er().Ty = eChrValEr &
'               eChrEmptyEr
If ZDtaEr_Sz = 0 Then Exit Function
Dim ChrEr() As ChrEr
    ChrEr = ZLnkInf_ChrEr
    
Dim LstDr() As LstDr
    LstDr = ZLnkInf_LstDr(ChrEr)

Dim O As TLstInf
    O.LstAdr = ZLnkInf_LstAdr(ChrEr, LstDr)
    O.LstSqv = ZLnkInf_LstSqv(LstDr)
ZB_ZLstInf = O
End Function

Private Function ZB_ZErMsg(LstAdr$()) As TErMsg()
Dim DtaEr() As TDtaEr
    DtaEr = ZDtaEr

Dim O() As TErMsg
    ReDim O(ZDtaEr_Sz - 1)
    Dim I&
    For I = 0 To ZDtaEr_Sz - 1
        Dim A$
        A = LstAdr$(I)
        O(I) = ZErMsg_One(DtaEr(I), A)
    Next
ZB_ZErMsg = O
End Function

Private Property Get ZDtaEr() As TDtaEr()
ZDtaEr = A_DtaEr.Ay
End Property

Private Property Get ZDtaEr_Sz%()
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

Private Property Get ZErCno_ErVal%()
ZErCno_ErVal = ZErCno("ErVal")
End Property

Private Property Get ZErCno_FldNmAy() As String()
ZErCno_FldNmAy = Split(ZErCno_Fld)
End Property

Private Property Get ZErCno_LstAdr%()
ZErCno_LstAdr = ZErCno("LstAdr")
End Property

Private Property Get ZErCno_Msg%()
ZErCno_Msg = ZErCno("Msg")
End Property

Private Property Get ZErCno_QDte%()
ZErCno_QDte = ZErCno("QDte")
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

Private Function ZErMsg_ChrCdNotFnd(Er As DE_ChrCdNotFnd, LstAdr) As TErMsg
Dim O As TErMsg
O.ErTxt = QErTxt.ChrCdNotFnd(Er.MsgDta)
With Er.ShwFld
    O.Adr = .Adr
    O.CharName = .CharName
    O.A_LstAdr = LstAdr
    O.CostEle = .CostEle
    O.CostGp = .CostGp
    O.FldNm = .FldNm
End With
ZErMsg_ChrCdNotFnd = O
End Function

Private Function ZErMsg_ChrEmpty(Er As DE_ChrEmpty, LstAdr) As TErMsg
Dim O As TErMsg
O.ErTxt = QErTxt.ChrEmpty()
With Er.ShwFld
    O.Ws = .Ws
    O.Adr = .Adr
    O.CharName = .CharName
    O.A_LstAdr = LstAdr
    O.CostEle = .CostEle
    O.CostGp = .CostGp
    O.FldNm = "Char" ' FldNm_OfChr(.CostGp, .CostEle, .CharName)
    O.Key = .Key
    O.OrgVal = .OrgVal
End With
ZErMsg_ChrEmpty = O
End Function

Private Function ZErMsg_ChrVal(Er As DE_ChrVal, LstAdr) As TErMsg
Dim O As TErMsg
O.ErTxt = QErTxt.ChrVal()
With Er.ShwFld
    O.Ws = .Ws
    O.Adr = .Adr
    O.CharName = .CharName
    O.A_LstAdr = LstAdr
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

Private Function ZErMsg_DifColCnt(Er As DE_DifColCnt) As TErMsg
Dim O As TErMsg
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

Private Function ZErMsg_DifHdCell(Er As DE_DifHdCell) As TErMsg
Dim O As TErMsg
O.ErTxt = QErTxt.DifHdCell(Er.MsgDta)
With Er.ShwFld
    O.Adr = .Adr
'    O.CharName = .CharName
'    O.ChrLnkAdr = ChrLnkAdr
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

Private Function ZErMsg_DifR1Formula(Er As DE_DifR1Formula) As TErMsg
Dim O As TErMsg
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

Private Function ZErMsg_DifVal(Er As DE_DifVal) As TErMsg
Dim O As TErMsg
O.ErTxt = QErTxt.DifVal(Er.MsgDta)
With Er.ShwFld
    O.ErVal = .ErVal
    O.Adr = .Adr
    O.FldNm = .FldNm
    O.Key = .Key
End With
ZErMsg_DifVal = O
End Function

Private Function ZErMsg_DupSku(Er As DE_DupSku) As TErMsg
Dim O As TErMsg
With Er
    O.ErTxt = QErTxt.DupSku(.MsgDta)
End With
With Er.ShwFld
    O.Ws = .Ws
    O.Key = .Key
    O.Adr = .Adr
    O.FldNm = .FldNm
End With
ZErMsg_DupSku = O
End Function

Private Property Get ZErMsg_HdSqv()
ZErMsg_HdSqv = Ay_HSqv(Split(ZErCno_Fld))
End Property

Private Function ZErMsg_NoOrgRow(Er As DE_NoOrgRow) As TErMsg
Dim O As TErMsg
O.ErTxt = QErTxt.NoOrgRow
With Er.ShwFld
    O.Key = .Key
    O.FldNm = .FldNm
    O.Adr = .Adr
End With
ZErMsg_NoOrgRow = O
End Function

Private Function ZErMsg_One(Er As TDtaEr, LstAdr$) As TErMsg
Dim O As TErMsg, T As TErTxt
Select Case Er.Ty
Case eDtaErTy.eChrValEr:       O = ZErMsg_ChrVal(Er.ChrVal, LstAdr)
Case eDtaErTy.eChrEmptyEr:     O = ZErMsg_ChrEmpty(Er.ChrEmpty, LstAdr)
Case eDtaErTy.eChrCdNotFndEr:  O = ZErMsg_ChrCdNotFnd(Er.ChrCdNotFnd, LstAdr)
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

Private Property Get ZErMsg_Sqv()
Dim ErMsg() As TErMsg
    ErMsg = B.ErMsg
    
Dim Fld$()
    Fld = ZErCno_FldNmAy

Dim NR%, NFld%
    NR = ZDtaEr_Sz
    NFld = Sz(Fld)
    
ReDim O(1 To NR, 1 To NFld)
    Dim J%, I%, K%
    For J = 0 To NR - 1
        I = 0
        With ErMsg(J)
            For K = 0 To NFld - 1
                Select Case Fld(K)
                Case "CharName":  I = I + 1: O(J + 1, I) = .CharName
                Case "LstAdr":    I = I + 1: O(J + 1, I) = .A_LstAdr
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
                Case Else: Stop
                End Select
            Next
        End With
    Next
ZErMsg_Sqv = O
End Property

Private Function ZErMsg_ValTy(Er As DE_ValTy) As TErMsg
Dim O As TErMsg
O.ErTxt = QErTxt.ValTy(Er.MsgDta)
With Er.ShwFld
    O.Ws = .Ws
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

Private Sub ZErWs_DoLnk_ColLstAdr_ToLstWs()
If ZLstSqv_NRow = 0 Then Exit Sub

Dim ErWs As Worksheet
    Set ErWs = B.ErWs

If IsNothing(ErWs) Then Exit Sub

Dim WrkWs As Worksheet
Dim OrgWs As Worksheet
Dim LstWs As Worksheet
    Set WrkWs = Src.Wrk.Ws
    Set OrgWs = Src.Org.Ws
    Set LstWs = Src_Wb.Sheets(LstWsNm)

Dim LstAdrCno%
    LstAdrCno = ZErCno_LstAdr

Dim J%
For J = 2 To Ws_LastRow(ErWs)
    Dim TarAdr$
        TarAdr = Ws_RC(ErWs, J, LstAdrCno).Value
    If Trim(TarAdr) = "" Then GoTo Nxt
        
    Dim TarRge As Range
        Set TarRge = LstWs.Range(TarAdr)
        
    Dim ErCell As Range
        Set ErCell = Ws_RC(ErWs, J, LstAdrCno)
        
    Cell_Lnk ErCell, TarRge        '<== Create HyperLink
Nxt:
Next
End Sub

Private Sub ZErWs_DoLnk_ColWrkAdr_ToSrcWs()
Dim ErWs As Worksheet
    Set ErWs = B.ErWs
Dim Wb As Workbook
    Set Wb = ErWs.Parent
Dim WrkWs As Worksheet
Dim OrgWs As Worksheet
    Set WrkWs = Wb.Sheets(WrkWsNm)
    Set OrgWs = Wb.Sheets(OrgWsNm)

Dim WrkAdrCno%
Dim WsCno%
    WrkAdrCno = ZErCno_WrkAdr
    WsCno = ZErCno_Ws

Dim J%
For J = 2 To Ws_LastRow(ErWs)
    Dim TarAdr$
        TarAdr = Ws_RC(ErWs, J, WrkAdrCno).Value
    
    If Trim(TarAdr) = "" Then GoTo Nxt
    
    Dim TarWsNm$
        TarWsNm = Ws_RC(ErWs, J, WsCno).Value
    
    Dim TarWs As Worksheet
        Select Case TarWsNm
        Case WrkWsNm, "": Set TarWs = WrkWs
        Case OrgWsNm: Set TarWs = OrgWs
        Case Else: Stop
        End Select
    
    Dim TarRge As Range
        Set TarRge = TarWs.Range(TarAdr)
    
    Dim Rge As Range
        Set Rge = Ws_RC(ErWs, J, WrkAdrCno)
    
    Cell_Lnk Rge, TarRge    '<=== Create HyperLink
Nxt:
Next
End Sub

Private Sub ZErWs_DoPut_ErMsg()
Dim Ws As Worksheet
    Set Ws = B.ErWs
Dim R2&
    R2 = ZDtaEr_Sz + 1
Dim C1%, C2%, C3%
    C1 = ZErCno_QDte
    C2 = ZErCno_ErVal
    C3 = ZErCno_Sku
Ws_CRR(Ws, C1, 2, R2).NumberFormat = "yyyy-mm-dd"
Ws_CRR(Ws, C2, 2, R2).NumberFormat = "@"
Ws_CRR(Ws, C3, 2, R2).NumberFormat = "@"

With Ws
    Cell_PutSqv .Range("A1"), ZErMsg_HdSqv
    Cell_PutSqv .Range("A2"), ZErMsg_Sqv
End With

Cell_Freeze Ws.Range("E2")
Ws_R(Ws, 1).AutoFilter
Ws.Columns.AutoFit
Ws_Zoom Ws, 85
Ws_C(Ws, ZErCno_ErVal).ColumnWidth = 30
Ws_C(Ws, ZErCno_Msg).ColumnWidth = 100
Ws_RR(Ws, 1, Ws_LastRow(Ws)).VerticalAlignment = xlVAlignCenter
Ws.Outline.SummaryColumn = xlSummaryOnLeft
End Sub

Private Function ZLnkInf_ChrEr() As ChrEr()
Dim DtaEr() As TDtaEr
    DtaEr = ZDtaEr
Dim U%
    U = UBound(DtaEr)

Dim LstAdrCno%
    LstAdrCno = ZErCno_LstAdr
Dim O() As ChrEr
    Dim J%
    For J = 0 To U
        Dim MM As TDtaEr
            MM = DtaEr(J)
        Select Case MM.Ty
        Case eChrValEr, eChrEmptyEr
        Case Else: GoTo Nxt
        End Select
        
        Dim M As ChrEr
            Select Case MM.Ty
            Case eChrValEr
                Dim M2 As DE_ChrVal
                    M2 = MM.ChrVal
                With M
                    .DtaErIdx = J
                    .Key = M2.ShwFld.Key
                    .WrkAdr = M2.ShwFld.Adr
                    .CharName = M2.ShwFld.CharName
                    .CharCode = M2.ChrDef.CharCode
                    .CostEle = M2.ShwFld.CostEle
                    .CostGp = M2.ShwFld.CostGp
                    .ErVal = M2.ShwFld.ErVal
                    .IsMulti = M2.ChrDef.IsMulti
                    .IsMust = M2.ChrDef.IsMust
                    Set .ValDic = M2.ChrDef.Dic_OfValNm_ToValCd
                    If IsNothing(.ValDic) Then Stop
                End With
            Case eChrEmptyEr
                Dim M1 As DE_ChrEmpty
                    M1 = MM.ChrEmpty
                With M
                    .DtaErIdx = J
                    .Key = M1.ShwFld.Key
                    .WrkAdr = M1.ShwFld.Adr
                    .CharName = M1.ShwFld.CharName
                    .CharCode = M1.ChrDef.CharCode
                    .CostEle = M1.ShwFld.CostEle
                    .CostGp = M1.ShwFld.CostGp
                    .IsMulti = M2.ChrDef.IsMulti
                    .IsMust = M2.ChrDef.IsMust
                    Set .ValDic = M1.ChrDef.Dic_OfValNm_ToValCd
                    If IsNothing(.ValDic) Then Stop
                End With
            Case Else
                Stop
            End Select
        If IsNothing(M.ValDic) Then Stop
        ZLnkInf_ChrErPush O, M
Nxt:
    Next
ZLnkInf_ChrEr = O
End Function

Private Function ZLnkInf_ChrErPush&(OAy() As ChrEr, M As ChrEr)
Dim N&
    N = ZLnkInf_ChrErSz(OAy)
ReDim Preserve OAy(N)
OAy(N) = M
End Function

Private Function ZLnkInf_ChrErSz&(ChrEr() As ChrEr)
On Error Resume Next
ZLnkInf_ChrErSz = UBound(ChrEr) + 1
End Function

Private Sub ZLnkInf_DoDmp_ChrEr(ChrEr() As ChrEr)
Dim J%
For J = 0 To ZLnkInf_ChrErSz(ChrEr) - 1
    With ChrEr(J)
        If .CharCode = "ZCASE_FABRICATION" Then Stop
        Debug.Print J, .WrkAdr,
    End With
Next
End Sub

Private Sub ZLnkInf_DoDmp_LstDr(LstDr() As LstDr)
Dim J&
For J = 0 To UBound(LstDr)
    With LstDr(J)
        If .IsFirst Then
            Debug.Print J, .ErAdr, .WrkAdr
        End If
    End With
Next
Stop
End Sub

Private Sub ZLnkInf_DoDmp_LstDr_ForNonBlankCharValName(Lst() As LstDr)
Dim J&, WithBlank As Boolean
For J = 0 To ZLnkInf_LstDrSz(Lst) - 1
    With Lst(J)
        If Trim(.CharValName) = "" Then Debug.Print J: WithBlank = True
    End With
Next
If WithBlank Then Stop
End Sub

Private Function ZLnkInf_LstAdr(ChrEr() As ChrEr, LstDr() As LstDr) As String()
'LstAdr is in the ErWs point to Ws-List-Column-ErAdr.  It is same size DtaEr
Dim C%
    C = ZLstCno_ErAdr
Dim O$()
    ReDim O(ZDtaEr_Sz - 1)
    Dim J&
    For J = 0 To ZLnkInf_LstDrSz(LstDr) - 1
        With LstDr(J)
            If .IsFirst Then
                Dim R&
                    R = 2 + J
                Dim LstAdr$
                    LstAdr = Ws_Adr(yWsMassUpd, R, C)
                Dim ChrErIdx%
                    ChrErIdx = LstDr(J).ChrErIdx
                Dim DtaErIdx%
                    DtaErIdx = ChrEr(ChrErIdx).DtaErIdx
                O(DtaErIdx) = LstAdr
            End If
        End With
    Next
ZLnkInf_LstAdr = O
End Function

Private Function ZLnkInf_LstDr(ChrEr() As ChrEr) As LstDr()
Dim NEr&
    NEr = ZLnkInf_ChrErSz(ChrEr)
If NEr = 0 Then Exit Function

Dim OLstDr() As LstDr
    Dim DrIdx&
        DrIdx = 0
        
    Dim ChrErIdx%
    For ChrErIdx = 0 To NEr - 1
        
        Dim A() As LstDr
            A = ZLnkInf_OneSetLstDr(ChrEr(ChrErIdx), ChrErIdx)
        ZLnkInf_LstDrPushAy OLstDr, A
Nxt:
    Next
ZLnkInf_LstDr = OLstDr
End Function

Private Sub ZLnkInf_LstDrPush(OAy() As LstDr, M As LstDr)
Dim N&
    N = ZLnkInf_LstDrSz(OAy)
ReDim Preserve OAy(N)
OAy(N) = M
End Sub

Private Sub ZLnkInf_LstDrPushAy(OAy() As LstDr, Ay() As LstDr)
Dim J%
For J = 0 To ZLnkInf_LstDrSz(Ay) - 1
    ZLnkInf_LstDrPush OAy, Ay(J)
Next
End Sub

Private Function ZLnkInf_LstDrSz&(Ay() As LstDr)
On Error Resume Next
ZLnkInf_LstDrSz = UBound(Ay) + 1
End Function

Private Function ZLnkInf_LstSqv(LstDr() As LstDr)
If ZLnkInf_LstDrSz(LstDr) = 0 Then Exit Function
Dim Fld$(), J&
    Fld = ZLstCno_FldNmAy

ReDim OSqv(1 To UBound(LstDr) + 2, 1 To Sz(Fld))
    For J = 0 To UBound(Fld)
        OSqv(1, J + 1) = Fld(J)
    Next
    Dim I%, K%
    For J = 0 To UBound(LstDr)
        K = 0
        With LstDr(J)
            If .IsFirst Then
                For I = 0 To UBound(Fld)
                    Select Case Fld(I)
                    Case "ErAdr":      K = K + 1: OSqv(2 + J, K) = .ErAdr
                    Case "WrkAdr":      K = K + 1: OSqv(2 + J, K) = .WrkAdr
                    Case "Pj":          K = K + 1: OSqv(2 + J, K) = .Key.Pj
                    Case "Sku":         K = K + 1: OSqv(2 + J, K) = .Key.Sku
                    Case "QDte":        K = K + 1: OSqv(2 + J, K) = .Key.QDte
                    Case "CostGp":      K = K + 1: OSqv(2 + J, K) = .CostGp
                    Case "CostEle":     K = K + 1: OSqv(2 + J, K) = .CostEle
                    Case "IsMust":      K = K + 1: OSqv(2 + J, K) = .IsMust
                    Case "IsMulti":     K = K + 1: OSqv(2 + J, K) = .IsMulti
                    Case "CharName":    K = K + 1: OSqv(2 + J, K) = .CharName
                    Case "ErVal":       K = K + 1: OSqv(2 + J, K) = .ErVal
                    Case "CharValName": K = K + 1: OSqv(2 + J, K) = .CharValName
                    Case Else: Stop
                    End Select
                Next
            Else
                For I = 0 To UBound(Fld)
                    Select Case Fld(I)
                    Case "ErAdr":      K = K + 1
                    Case "WrkAdr":      K = K + 1
                    Case "Pj":          K = K + 1
                    Case "Sku":         K = K + 1
                    Case "QDte":        K = K + 1
                    Case "CostGp":      K = K + 1
                    Case "CostEle":     K = K + 1
                    Case "IsMust":      K = K + 1
                    Case "IsMulti":     K = K + 1
                    Case "CharName":    K = K + 1
                    Case "ErVal":       K = K + 1
                    Case "CharValName": K = K + 1: OSqv(2 + J, K) = .CharValName
                    Case Else: Stop
                    End Select
                Next
            End If
        End With
    Next
ZLnkInf_LstSqv = OSqv
End Function

Private Function ZLnkInf_OneSetLstDr(C As ChrEr, ChrErIdx%) As LstDr()
If C.ValDic.Count = 0 Then Exit Function

Dim U%
    U = C.ValDic.Count - 1

Dim ValNm()
    ValNm = C.ValDic.Keys

Dim Cno%
    Cno = ZErCno_LstAdr
    
Dim O() As LstDr
    ReDim O(U)
    Dim J%
    For J = 0 To U
        With O(J)
            If J = 0 Then
                .ChrErIdx = ChrErIdx
                .WrkAdr = C.WrkAdr
                .ErAdr = Ws_Adr(yWsMassUpd, C.DtaErIdx + 2, Cno)
                .IsFirst = True
                .CharName = C.CharName
                .CostGp = C.CostGp
                .CharValName = C.CharName
                .CostEle = C.CostEle
                .IsMulti = C.IsMulti
                .IsMust = C.IsMust
                .ErVal = C.ErVal
                .Key = C.Key
            End If
            .CharValName = ValNm(J)
        End With
    Next
ZLnkInf_OneSetLstDr = O
End Function

Private Function ZLstCno%(FldNm$)
Dim O%
O = Ay_Idx(ZLstCno_FldNmAy, FldNm) + 1
If O = 0 Then Stop
ZLstCno = O
End Function

Private Property Get ZLstCno_CharValName%()
ZLstCno_CharValName = ZLstCno("CharValName")
End Property

Private Property Get ZLstCno_ErAdr%()
ZLstCno_ErAdr = ZLstCno("ErAdr")
End Property

Private Property Get ZLstCno_ErVal%()
ZLstCno_ErVal = ZLstCno("ErVal")
End Property

Private Property Get ZLstCno_FldNmAy() As String()
ZLstCno_FldNmAy = Split(ZLstCno_Fld)
End Property

Private Property Get ZLstCno_QDte%()
ZLstCno_QDte = ZLstCno("QDte")
End Property

Private Property Get ZLstCno_Sku%()
ZLstCno_Sku = ZLstCno("Sku")
End Property

Private Property Get ZLstCno_WrkAdr%()
ZLstCno_WrkAdr = ZLstCno("WrkAdr")
End Property

Private Function ZLstSqv_NRow&()
If Not IsEmpty(B.LstSqv) Then
    ZLstSqv_NRow = UBound(B.LstSqv, 1)
End If
End Function

Private Sub ZLstWs_DoCrt_ChgSelectionMacro()
Ws_Crt_EvtMth_CallingFn B.LstWs, "SelectionChange", "jjMassUpd", "Macro_LstWs"
End Sub

Private Function ZLstWs_DoCrt_EmptyIfNeeded(Wb As Workbook) As Worksheet
If ZLstSqv_NRow > 0 Then
   Set ZLstWs_DoCrt_EmptyIfNeeded = Wb_AddWs_AtEnd(Wb, LstWsNm)
End If
End Function

Private Sub ZLstWs_DoLnk_ColLnk_ToErWs()
Dim LstWs As Worksheet
    Set LstWs = B.LstWs
If IsNothing(LstWs) Then Exit Sub

Dim ErWs As Worksheet
    Set ErWs = B.ErWs
    
Dim ErAdrCno%
    ErAdrCno = ZLstCno_ErAdr

Dim Sqv
    Sqv = Ws_Sqv_ByA1ToLastCell_withR1(LstWs)

Dim R&
Dim LnkCell As Range      ' The lnk-cell in LstWs requires a link
Dim LnkTar As Range       ' The lnk-cell's target, which is in LnkWs

For R = 2 To UBound(Sqv, 1)
    If IsEmpty(Sqv(R, ErAdrCno)) Then GoTo Nxt

    Set LnkCell = LstWs.Cells(R, ErAdrCno)
    Set LnkTar = ErWs.Range(LnkCell.Value)
    
    Cell_Lnk LnkCell, LnkTar
Nxt:
Next
End Sub

Private Sub ZLstWs_DoLnk_ColWrk_ToWrkWs()
'Create Link in LstWs to 2-col:(WrkAdr LnkAdr)
Dim LstWs As Worksheet
    Set LstWs = B.LstWs
If IsNothing(LstWs) Then Exit Sub

Dim Wb As Workbook
    Set Wb = LstWs.Parent
Dim WrkWs As Worksheet
    Set WrkWs = Wb.Sheets(WrkWsNm)
Dim WrkAdrCno%
    WrkAdrCno = ZLstCno_WrkAdr

Dim Sqv
    Sqv = Ws_Sqv_ByA1ToLastCell_withR1(LstWs)

Dim R&
Dim WrkCell As Range      ' The wrk-cell in LstWs requires a link
Dim WrkTar As Range       ' The wrk-cell's target, which is in WrkWs

For R = 2 To UBound(Sqv, 1)
    If IsEmpty(Sqv(R, WrkAdrCno)) Then GoTo Nxt
    
    Set WrkCell = LstWs.Cells(R, WrkAdrCno)
    Set WrkTar = WrkWs.Range(WrkCell.Value)
    
    Cell_Lnk WrkCell, WrkTar
Nxt:
Next
End Sub

Private Sub ZLstWs_DoPut_LstSqv()
Dim LstWs As Worksheet
    Set LstWs = B.LstWs
If IsNothing(LstWs) Then Exit Sub
Dim N&
    N = ZLstSqv_NRow
If N = 0 Then Exit Sub
ZLstWs_FmtCol LstWs, N, "@", ZLstCno_CharValName
ZLstWs_FmtCol LstWs, N, "@", ZLstCno_Sku
ZLstWs_FmtCol LstWs, N, "@", ZLstCno_ErVal
ZLstWs_FmtCol LstWs, N, "yyyy-mm-dd", ZLstCno_QDte
Cell_PutSqv LstWs.Range("A1"), B.LstSqv

'====
Cell_Freeze LstWs.Range("D2")
Ws_R(LstWs, 1).AutoFilter
LstWs.Columns.AutoFit
Ws_Zoom LstWs, 85
Ws_RR(LstWs, 1, Ws_LastRow(LstWs)).VerticalAlignment = xlVAlignCenter
Ws_C(LstWs, ZLstCno_ErVal).ColumnWidth = 30
Ws_C(LstWs, ZLstCno_CharValName).ColumnWidth = 100
End Sub

Private Sub ZLstWs_FmtCol(LstWs As Worksheet, LstWs_NRow&, NbrFmtStr$, Cno%)
Ws_CRR(LstWs, Cno, 2, LstWs_NRow + 1).NumberFormat = NbrFmtStr$
End Sub

Private Sub ZSrcWs_DoLnk_ErCell_ToErWs()
Dim ErWs As Worksheet
    Set ErWs = B.ErWs
Dim Wrk As Worksheet
Dim Org As Worksheet
    Dim Wb As Workbook
    Set Wb = ErWs.Parent
    Set Wrk = Wb.Sheets(WrkWsNm)
    Set Org = Wb.Sheets(OrgWsNm)

'Add a lnk to SrcCell which will jmp to ErCell
Wrk.Hyperlinks.Delete   'Assume there is only error link. So clear all should be OK
Org.Hyperlinks.Delete   'Assume
Wrk.ListObjects(1).DataBodyRange.Font.ColorIndex = xlAutomatic
Org.ListObjects(1).DataBodyRange.Font.ColorIndex = xlAutomatic

Dim C%                  ' The AdrCno in ErWrk
    C = ZErCno_WrkAdr

Dim J&
For J = 0 To ZDtaEr_Sz - 1
    Dim Er As TErMsg
        Er = B.ErMsg(J)
    
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
