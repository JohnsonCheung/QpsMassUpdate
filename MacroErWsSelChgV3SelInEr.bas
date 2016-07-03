Attribute VB_Name = "MacroErWsSelChgV3SelInEr"
Option Explicit
Private A_Cell As Range
Private Type VBar
    R1 As Long
    R2 As Long
    C As Integer
End Type

Sub ErWs_WsSelChg_V3SelInEr(Cell As Range)
Set A_Cell = Cell
If ZIsCurCell_InSide_CorrVBar Then ZSel_DoBld: Exit Sub
If ZIsCurCell_OutSide_SelCol Then ZSel_DoClr: Exit Sub
If ZIsCurCell_InSide_SelValVBar Then
    If ZIsMulti Then
        ZSelCell_DoToggle
    Else
        ZSelCell_DoSelect
    End If
    ZCorr_DoSetCellVal
    ZTar_DoSetCellVal
    Exit Sub
End If
If ZIsCurCell_InSide_SelVbar Then ZSel_DoBld: Exit Sub
End Sub

Private Sub ZCell_DoHighlight(Cell As Range)
Rge_RC(Cell, 1, 1).Interior.Color = rgbYellow
End Sub

Private Sub ZCell_DoSetVal(Cell As Range, ValStr$)
With Cell
    With .Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    .Font.ColorIndex = xlAutomatic
    .Font.TintAndShade = 0
    .Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone
    .NumberFormat = "@"
    .Value = ValStr
End With
End Sub

Private Sub ZCell_DoUnlight(Cell As Range)
Rge_RC(Cell, 1, 1).Interior.Pattern = xlNone
End Sub

Private Property Get ZChrNm_CharValNameAy() As String()
ZChrNm_CharValNameAy = ChrDefInf.ChrNm_ChrValNmAy(ZChrNm_CurVal)
End Property

Private Property Get ZChrNm_Cno%()
ZChrNm_Cno = ZZFndCno("CharName")
End Property

Private Property Get ZChrNm_CurVal$()
ZChrNm_CurVal = Ws_RC(ZWs, ZR, ZChrNm_Cno).Value
End Property

Private Function ZChrValNm_Cell1() As Range
Dim C%
    C = ZSel_Cno
Dim Rge0 As Range
    Set Rge0 = Ws_RC(ZWs, 1, C)
    
Dim A As Range 'Below cell
    Set A = Rge_RC(Rge0, 2, 1)
    
Dim Rge1 As Range
    If IsEmpty(A.Value) Then
        Set Rge1 = Rge0.End(xlDown)
    Else
        Set Rge1 = A
    End If
Set ZChrValNm_Cell1 = Rge1
End Function

Private Function ZChrValNm_Cell2() As Range
Dim Rge0 As Range, Rge1 As Range
    Set Rge0 = ZChrValNm_Cell1
    Set Rge1 = Rge0.End(xlDown)
Set ZChrValNm_Cell2 = Rge1
End Function

Private Sub ZChrValNm_DoHighlight_FmCorrVal()
Dim C As Range
    Set C = ZCorr_Cell
If IsNothing(C) Then Exit Sub
Dim A$
    A = ZCorr_Cell.Value
Dim B$()
    B = Split(A, vbLf)
Dim R As Range
    Set R = ZChrValNm_Rge
If R.Count = 0 Then Exit Sub
Dim Cell As Range
For Each Cell In R
    Dim CellVal$
        CellVal = Cell.Value
    If Ay_Has(B, CellVal) Then ZCell_DoHighlight Cell '<====
Next
End Sub

Private Sub ZChrValNm_DoUnlight()
Dim Ay() As Range, J%
ZChrValNm_Rge.Interior.Pattern = xlNone
End Sub

Private Property Get ZChrValNm_Rge() As Range
Dim B As VBar
B = ZSel_ValVBar
Set ZChrValNm_Rge = Ws_CRR(ZWs, B.C, B.R1, B.R2)
End Property

Private Function ZCorr_Cell() As Range
If Not ZIsCurCell_InSide_SelValVBar Then
    If Not ZIsCurCell_InSide_CorrVBar Then
        Exit Function
    End If
End If

Dim R1&
    R1 = ZChrValNm_Cell1.Row

Set ZCorr_Cell = Ws_RC(ZWs, R1, ZCorr_Cno)
End Function

Private Property Get ZCorr_CellAdr$()
ZCorr_CellAdr = ZCorr_Cell.Address
End Property

Private Function ZCorr_Cno%()
ZCorr_Cno = ZZFndCno("Correction")
End Function

Private Sub ZCorr_DoSetCellVal()
Dim A$
    A = ZSelected_NewValStr
ZCell_DoSetVal ZCorr_Cell, A
End Sub

Private Sub ZCorr_DoSetColAsTxt()
Ws_C(ZWs, ZCorr_Cno).NumberFormat = "@"
End Sub

Private Function ZCorr_VBar() As VBar
Dim O As VBar
O.C = ZCorr_Cno
O.R1 = 2
O.R2 = ZR2
ZCorr_VBar = O
End Function

Private Sub ZCurCell_DoScrollToTop()
ActiveWindow.ScrollRow = A_Cell.Row
End Sub

Private Property Get ZCurCell_PossibleSelectionSqv()
Dim A$()
    A = ZChrNm_CharValNameAy
Dim N&
    N = Sz(A)
If N = 0 Then Exit Property
ReDim O(1 To N, 1 To 1)
    Dim J%
    For J = 1 To N
        O(J, 1) = A(J - 1)
    Next
ZCurCell_PossibleSelectionSqv = O
End Property

Private Function ZIsCell_Highlight(Cell As Range) As Boolean
ZIsCell_Highlight = Cell.Interior.Color = rgbYellow
End Function

Private Property Get ZIsCurCell_InSide_CorrVBar() As Boolean
ZIsCurCell_InSide_CorrVBar = ZIsCurCell_InSide_VBar(ZCorr_VBar)
End Property

Private Property Get ZIsCurCell_InSide_SelValVBar() As Boolean
ZIsCurCell_InSide_SelValVBar = ZIsCurCell_InSide_VBar(ZSel_ValVBar)
End Property

Private Property Get ZIsCurCell_InSide_SelVbar() As Boolean
ZIsCurCell_InSide_SelVbar = ZIsCurCell_InSide_VBar(ZSel_VBar)
End Property

Private Property Get ZIsCurCell_InSide_VBar(B As VBar) As Boolean
ZIsCurCell_InSide_VBar = Not ZIsCurCell_OutSide_VBar(B)
End Property

Private Property Get ZIsCurCell_OutSide_SelCol() As Boolean
ZIsCurCell_OutSide_SelCol = ZIsCurCell_OutSide_VBar(ZSel_ColVBar)
End Property

Private Function ZIsCurCell_OutSide_VBar(B As VBar) As Boolean
Dim R&, C%
    R = A_Cell.Row
    C = A_Cell.Column
ZIsCurCell_OutSide_VBar = True
If B.R1 > R Or R > B.R2 Then Exit Function
If C <> B.C Then Exit Function
ZIsCurCell_OutSide_VBar = False
End Function

Private Function ZIsMulti() As Boolean
ZIsMulti = ZMulti_Cell.Value = "Multi"
End Function

Private Function ZIsMust() As Boolean
ZIsMust = ZMust_Cell.Value = "Must"
End Function

Private Function ZIsSingle() As Boolean
ZIsSingle = Not ZIsMulti
End Function

Private Function ZMulti_Cell() As Range
If Not ZIsCurCell_InSide_SelValVBar Then Exit Function
Dim R&
    R = ZChrValNm_Cell1.Row
Set ZMulti_Cell = A_Cell.Worksheet.Cells(R, ZMulti_Cno)
End Function

Private Property Get ZMulti_Cno%()
ZMulti_Cno = ZZFndCno("Multi")
End Property

Private Function ZMust_Cell() As Range
Set ZMust_Cell = A_Cell.Worksheet.Cells(ZR, ZMust_Cno)
End Function

Private Property Get ZMust_Cno%()
ZMust_Cno = ZZFndCno("Must")
End Property

Private Property Get ZR2&()
ZR2 = Ws_RC(ZWs, 1, 1).End(xlDown).Row
End Property

Private Property Get ZR&()
ZR = A_Cell.Row
End Property

Private Sub ZSelCell_DoHighlight()
ZCell_DoHighlight A_Cell
End Sub

Private Sub ZSelCell_DoSelect()
ZChrValNm_DoUnlight
ZSelCell_DoHighlight
End Sub

Private Sub ZSelCell_DoToggle()
If ZSelCell_IsHighlight Then
    ZSelCell_DoUnlight
Else
    ZSelCell_DoHighlight
End If
End Sub

Private Sub ZSelCell_DoUnlight()
ZCell_DoUnlight A_Cell
End Sub

Private Function ZSelCell_IsHighlight() As Boolean
ZSelCell_IsHighlight = ZIsCell_Highlight(A_Cell)
End Function

Private Function ZSel_Cno%()
ZSel_Cno = ZZFndCno("Selection")
End Function

Private Function ZSel_ColVBar() As VBar
Dim O As VBar
O.C = ZSel_Cno
O.R1 = 2
O.R2 = Ws_MaxRno(ZWs)
ZSel_ColVBar = O
End Function

Private Sub ZSel_DoBld()
ZSel_DoClr
   
Dim Sqv
    Sqv = ZCurCell_PossibleSelectionSqv
If IsEmpty(Sqv) Then Exit Sub

Dim N%
    N = UBound(Sqv, 1)
    
Dim C%, R1&, R2&
    C = ZSel_Cno
    R1 = ZR
    R2 = R1 + N - 1
Dim R As Range
    Set R = Ws_CRR(ZWs, C, R1, R2)
    R.NumberFormat = "@"
    R.VerticalAlignment = XlVAlign.xlVAlignCenter
R.Value = Sqv
ZChrValNm_DoHighlight_FmCorrVal
End Sub

Private Sub ZSel_DoClr()
Dim C%
    C = ZSel_Cno
Dim R1&, R2&
    R1 = 2
    R2 = Ws_MaxRno(ZWs)
With Ws_CRR(ZWs, C, R1, R2) ' Don't use .Clear
    .Value = Empty
    With .Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End With
End Sub

Private Sub ZSel_DoSetColAsTxt()
Ws_C(ZWs, ZSel_Cno).NumberFormat = "@"
End Sub

Private Function ZSel_VBar() As VBar
Dim O As VBar
O.R1 = 2
O.R2 = ZR2
O.C = ZSel_Cno
ZSel_VBar = O
End Function

Private Function ZSel_ValVBar() As VBar
Dim R1&, R2&
    R1 = ZChrValNm_Cell1.Row
    R2 = ZChrValNm_Cell2.Row
Dim O As VBar
O.R1 = R1
O.R2 = R2
O.C = ZSel_Cno
ZSel_ValVBar = O
End Function

Private Property Get ZSelected_CharValNameAy() As String()
Dim A As Range
    Set A = ZChrValNm_Rge
If A.Count = 0 Then Exit Property
Dim O$()
    Dim R As Range
    For Each R In A
        If ZIsCell_Highlight(R) Then Push O, R.Value
    Next
ZSelected_CharValNameAy = O
End Property

Private Function ZSelected_NewValStr$()
ZSelected_NewValStr = Join(ZSelected_CharValNameAy, vbLf)
End Function

Private Function ZTar_Cell() As Range
Set ZTar_Cell = ZTar_Ws.Range(ZWrkAdr_CellVal)
End Function

Private Function ZTar_CellVal_Cur$()
ZTar_CellVal_Cur = ZTar_Cell.Value
End Function

Private Sub ZTar_Cell_DoResetColor()
ZTar_Cell.Font.ColorIndex = xlAutomatic
End Sub

Private Sub ZTar_Cell_DoSetRed()
ZTar_Cell.Font.Color = rgbRed
End Sub

Private Sub ZTar_DoSetCellVal()
If Not ZTar_WsExist Then Exit Sub
ZCell_DoSetVal ZTar_Cell, ZSelected_NewValStr
End Sub

Private Function ZTar_Ws() As Worksheet
Set ZTar_Ws = A_Cell.Worksheet.Parent.Sheets("Working")
End Function

Private Property Get ZTar_WsExist() As Boolean
ZTar_WsExist = Wb_IsWs(Ws_Wb(A_Cell.Worksheet), "Working")
End Property

Private Function ZWrkAdr_Cell() As Range
Set ZWrkAdr_Cell = A_Cell.Worksheet.Cells(ZChrValNm_Cell1.Row, ZWrkAdr_Cno)
End Function

Private Function ZWrkAdr_CellVal$()
ZWrkAdr_CellVal = ZWrkAdr_Cell.Value
End Function

Private Function ZWrkAdr_Cno%()
ZWrkAdr_Cno = ZZFndCno("WrkAdr")
End Function

Private Function ZWs() As Worksheet
Set ZWs = A_Cell.Worksheet
End Function

Private Function ZZFndCno%(ColNm$)
Dim Ws As Worksheet
    Set Ws = ZWs

Dim C%
    C = Ws_RC(Ws, 1, 1).End(xlToRight).Column

Dim Sqv
    Sqv = Ws_RCC(Ws, 1, 1, C).Value

Dim O%
    For O = 1 To UBound(Sqv, 2)
        If Sqv(1, O) = ColNm Then ZZFndCno = O: Exit Function
    Next
Stop
End Function
