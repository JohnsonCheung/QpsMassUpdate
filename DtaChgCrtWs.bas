Attribute VB_Name = "DtaChgCrtWs"
Option Explicit
Option Compare Text
'R C Pj Sku QDte CostGp CostEle CharName OrgVal WrkVal Fld Type
'1 2 3  4   5    6      7       8        9      10     11  12
'X X X      X                            O      N      6*  PjQ      // 6-Fld: (QpsFormFileName / Date / Size / Rate * 3
'X X X      X                            O      N      4*  One      // 4-Fld: ProtoCst ProtoRmk ToolCst ToolRmk
'X X X  X   X                            O      N      8*  Sku      // 8-Fld: PotentialQty Sku 3-DrX*(USD+HKD)
'X X X  X   X    X      X                O      N          CstVal
'X X X  X   X    X      X                O      N          CstRmk
'X X X  X   X    X      X       X        O      N          Chr
Dim X_FldNmAy$()
Dim X_Ws As Worksheet
Dim X_R1&
Dim X_R2&
Const FldLst$ = "Pj Sku QDte FldNm CostGp CostEle CharName OrgVal WrkVal Adr"

Sub DtaChg_DoCrt_Ws_and_PaintWrkWs(Ay() As TDtaChg, Wb As Workbook)
If IsNothing(Wb) Then Exit Sub
Dim NR&
    NR = DtaChg_Sz(Ay)
If NR = 0 Then Exit Sub

X_FldNmAy = Split(FldLst)
Set X_Ws = Wb_AddWs_AtEnd(Wb, DtaChgWsNm, True)
X_R1 = 2
X_R2 = NR + 1

ZColRge("QDte").NumberFormat = "yyyy-mm-dd"
ZColRge("SKU").NumberFormat = "@"
ZColRge("OrgVal").HorizontalAlignment = XlHAlign.xlHAlignLeft
ZColRge("WrkVal").Interior.Color = rgbYellow
ZColRge("WrkVal").HorizontalAlignment = XlHAlign.xlHAlignLeft
Z2Col("OrgVal", "WrkVal").BorderAround XlLineStyle.xlContinuous, xlMedium
'==== Put Data to RsltWs ========
Cell_PutSqv X_Ws.Range("A1"), ZHdSqv
Cell_PutSqv X_Ws.Range("A2"), ZSqv(Ay)
'==== Put Data to RsltWs ========
'ZWs_DoFreezeAt     Cell-A2
'ZWs_DoAutoFilter   Row-1
'ZWs_DoAutoFit      Columns-A1
'ZWs_DoVAlignCentre AllCells-A1-to-LastCell
'ZWs_DoLnk          Col-Adr
Cell_Freeze X_Ws.Range("A2")
X_Ws.Range("A1").AutoFilter
X_Ws.Columns.AutoFit
Ws_A1_To_LastCell(X_Ws).VerticalAlignment = XlVAlign.xlVAlignCenter
ZLnkAdrCol
Ws_Zoom X_Ws, 85

Dim DtaChgWs As Worksheet
    Set DtaChgWs = Wb.Sheets(DtaChgWsNm)

Dim WrkWs As Worksheet
    Set WrkWs = Wb.Sheets(WrkWsNm)
    
Dim Y As YellowAdr
    Y = DtaChg_YellowAdr(Ay)

FmtFilter_DoRestore WrkWs
FmtFilter_DoSetYellow WrkWs, Y
FmtColor_DoPaint_Yellow WrkWs, Y

ZLnkWrkWs_ToDtaChgWs Y, WrkWs, DtaChgWs
End Sub

Private Function Z2Col(ColNm1$, ColNm2$) As Range
Set Z2Col = Ws_RCRC(X_Ws, X_R1, ZCno(ColNm1), X_R2, ZCno(ColNm2))
End Function

Private Function ZCno%(FldNm$)
Dim O%
O = Ay_Idx(X_FldNmAy, FldNm) + 1
If O = 0 Then Stop
ZCno = O
End Function

Private Function ZColRge(ColNm$) As Range
Set ZColRge = Ws_CRR(X_Ws, ZCno(ColNm), X_R1, X_R2)
End Function

Private Function ZHdSqv()
'Public Const C_DcWs_FldLst = "Adr Pj Sku QDte FldNm CostGp CostEle CharName OrgVal WrkVal"
ZHdSqv = Ay_HSqv(Split(FldLst))
End Function

Private Sub ZLnkAdrCol()
Dim R As Range, Ws As Worksheet
Set Ws = Ws_Wb(X_Ws).Sheets(WrkWsNm)
For Each R In ZColRge("Adr")
    Cell_Lnk R, Ws.Range(R.Value)
Next
End Sub

Private Sub ZLnkWrkWs_ToDtaChgWs(YellowAdr As YellowAdr, WrkWs As Worksheet, DcWs As Worksheet)
Dim N&, J&, Src As Range, Tar As Range
N = Vb_Ay.Sz(YellowAdr.C)
If N = 0 Then Exit Sub
For J = 0 To N - 1
    Set Src = WrkWs.Cells(YellowAdr.R(J), YellowAdr.C(J))
    Set Tar = DcWs.Cells(J + 2, "I")
    If Trim(Src.Value) <> "" Then
        Cell_Lnk Src, Tar
        Src.WrapText = True
    End If
Next
End Sub

Private Function ZSqv(DtaChg() As TDtaChg)
Dim NR&, R&, D As TDtaChg, C%
Dim J%, K%
NR = UBound(DtaChg) + 1
ReDim O(1 To NR, 1 To UBound(X_FldNmAy) + 1)
For R = 1 To NR
    C = 0
    D = DtaChg(R - 1)
    For J = 0 To UB(X_FldNmAy)
        Select Case X_FldNmAy(J)
        Case "Pj":       C = C + 1: O(R, C) = D.Key.Pj
        Case "Sku":      C = C + 1: O(R, C) = D.Key.Sku
        Case "QDte":     C = C + 1: O(R, C) = D.Key.QDte
        Case "WrkVal":   C = C + 1: O(R, C) = D.WrkVal
        Case "OrgVal":   C = C + 1: O(R, C) = D.OrgVal
        Case "CostGp":   C = C + 1: O(R, C) = D.CostGp
        Case "CostEle":  C = C + 1: O(R, C) = D.CostEle
        Case "CharName": C = C + 1: O(R, C) = D.CharName
        Case "CostGp":   C = C + 1: O(R, C) = D.CostGp
        Case "FldNm":    C = C + 1: O(R, C) = D.FldNm
        Case "Adr":      C = C + 1: O(R, C) = Fct.SrcSqvAdr(D.Key.Rno, D.Cno)
        Case Else: Stop
        End Select
    Next
Next
ZSqv = O
End Function
