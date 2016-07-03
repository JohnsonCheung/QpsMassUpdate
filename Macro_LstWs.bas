Attribute VB_Name = "Macro_LstWs"
Option Explicit
Private A_Cell As Range

Sub SelectionChange(Cell As Range)
Set A_Cell = Cell
If ZIsCell_OutSideRange Then Exit Sub
If ZIsMulti Then
    ZDoToggle_Cell
Else
    ZDoSelect_Cell
End If
ZDoSet_TheCellVal
'------------------------------

ZVdt_Again           'Vdt again if any error.  Shw the message @ col-after-CharValName and first-row
End Sub

Private Property Get YAdr_IsMultiCell$()
YAdr_IsMultiCell = ZIsMultiCell.Address
End Property

Private Property Get YAdr_LeftCell$()
YAdr_LeftCell$ = ZLeftCell_UnderSku.Address
End Property

Private Property Get YAdr_TheCell$()
YAdr_TheCell = ZWrkCell.Address
End Property

Private Property Get YAdr_WrkAdrCell$()
YAdr_WrkAdrCell = ZWrkAdrCell.Address
End Property

Private Property Get YCno_CharValName%()
YCno_CharValName = ZCharValNameCno
End Property

Private Property Get YCno_WrkAdr%()
YCno_WrkAdr = ZWrkAdrCno
End Property

Private Property Get YR1&()
YR1 = ZR1
End Property

Private Property Get YR2&()
YR2 = ZR2
End Property

Private Property Get YRLast&()
YRLast = ZRLast
End Property

Private Property Get YTheCellVal_Cur$()
YTheCellVal_Cur = ZWrkCell_CurVal
End Property

Private Property Get YTheCellVal_New$()
YTheCellVal_New = ZWrkCell_NewVal
End Property

Private Function ZAll_CharValName_Rge() As Range
Set ZAll_CharValName_Rge = Ws_CRR(A_Cell.Worksheet, A_Cell.Column, ZR1, ZR2)
End Function

Private Function ZAll_CharValName_RgeAdr$()
ZAll_CharValName_RgeAdr$ = ZAll_CharValName_Rge.Address
End Function

Private Function ZAll_SelectedCell() As Range()
Dim O() As Range, R As Range
For Each R In ZAll_CharValName_Rge
    If ZIsHigh_A_Cell(R) Then Push O, R
Next
ZAll_SelectedCell = O
End Function

Private Function ZAll_Selected_CharValName() As String()
Dim Ay() As Range, O$(), J%
Ay = ZAll_SelectedCell
For J = 0 To UB(Ay)
    Push O, Ay(J).Value
Next
ZAll_Selected_CharValName = O
End Function

Private Function ZCharValNameCno%()
ZCharValNameCno% = ZFndCno("CharValName")
End Function

Private Sub ZDoHighlight_A_Cell(Cell As Range)
Rge_RC(Cell, 1, 1).Interior.Color = rgbYellow
End Sub

Private Sub ZDoHighlight_Cell()
ZDoHighlight_A_Cell A_Cell
End Sub

Private Sub ZDoReset_ErMsg()
ZErMsgCell.Clear
End Sub

Private Sub ZDoReset_TheCellColor()
ZWrkCell.Font.ColorIndex = xlAutomatic
End Sub

Private Sub ZDoSelect_Cell()
ZDoUnLight_All
ZDoHighlight_Cell
End Sub

Private Sub ZDoSet_ErMsg()
ZErMsgCell.Value = ZErMsg
End Sub

Private Sub ZDoSet_TheCellToRed()
ZWrkCell.Font.Color = rgbRed
End Sub

Private Sub ZDoSet_TheCellVal()
With ZWrkCell
    .Interior.ColorIndex = xlNone
    .NumberFormat = "@"
    .Value = ZWrkCell_NewVal
End With
End Sub

Private Sub ZDoToggle_Cell()
If ZIsHigh_Cell Then
    ZDoUnlight_Cell
Else
    ZDoHighlight_Cell
End If
End Sub

Private Sub ZDoUnLight_All()
Dim Ay() As Range, J%
If True Then
    Ay = ZAll_SelectedCell
    For J = 0 To UB(Ay)
        ZDoUnlight_A_Cell Ay(J)
    Next
Else
    ZAll_CharValName_Rge.Interior.Pattern = xlNone
End If
End Sub

Private Sub ZDoUnlight_A_Cell(Cell As Range)
Rge_RC(Cell, 1, 1).Interior.Pattern = xlNone
End Sub

Private Sub ZDoUnlight_Cell()
ZDoUnlight_A_Cell A_Cell
End Sub

Private Function ZErMsg$()
Dim O$()
If ZIsEr_Must Then
    Push O, "This char must be entered"
Else
    If ZIsEr_MultiValEntered_ForSingle Then
        Push O, "Multiple values has been entered, but the Characteristics only allow single value"
    End If
    If ZIsEr_InvdtValEntered Then
        Push O, Fmt_QQ("This value entered ? are invalid", ZInvdtValEntered_Str)
    End If
End If
ZErMsg = Join(O, vbLf)
End Function

Private Function ZErMsgCell() As Range
Set ZErMsgCell = A_Cell.Worksheet.Cells(ZR1, A_Cell.Column + 1)
End Function

Private Function ZFndCno%(ColNm$)
Dim J%, Sqv
Sqv = Ws_RCC(A_Cell.Worksheet, 1, 1, 20).Value
For J = 1 To 100
    If Sqv(1, J) = ColNm Then ZFndCno = J: Exit Function
Next
End Function

Private Function ZInvdtValEntered_Str$()
ZInvdtValEntered_Str = Join(Ay_Quote(ZInvdtValEntered_StrAy, "[]"), " ")
End Function

Private Function ZInvdtValEntered_StrAy() As String()
Dim Ay$(), Vdt$(), O$(), J%
Ay = Split(ZWrkCell_CurVal, vbLf)
Vdt = ZVdtVal
For J = 0 To UB(Ay)
    If Ay_Idx(Vdt, Ay(J)) < 0 Then Push O, Ay(J)
Next
ZInvdtValEntered_StrAy = O
End Function

Private Function ZIsCell_OutSideRange() As Boolean
ZIsCell_OutSideRange = True
If A_Cell.Row = 1 Then Exit Function
If A_Cell.Column <> ZCharValNameCno Then Exit Function
If A_Cell.Row > ZRLast Then Exit Function

ZIsCell_OutSideRange = False
End Function

Private Function ZIsEr_InvdtValEntered() As Boolean
ZIsEr_InvdtValEntered = Sz(ZInvdtValEntered_StrAy) > 0
End Function

Private Function ZIsEr_MultiValEntered_ForSingle() As Boolean
If Not ZIsSingle Then Exit Function
If Not ZIsMultiValEntered Then Exit Function
ZIsEr_MultiValEntered_ForSingle = True
End Function

Private Function ZIsEr_Must() As Boolean
If Not ZIsMust Then Exit Function
If Not ZIsNoSelection Then Exit Function
ZIsEr_Must = True
End Function

Private Function ZIsHigh_A_Cell(Cell As Range) As Boolean
ZIsHigh_A_Cell = Cell.Interior.Color = rgbYellow
End Function

Private Function ZIsHigh_Cell() As Boolean
ZIsHigh_Cell = ZIsHigh_A_Cell(A_Cell)
End Function

Private Function ZIsMulti() As Boolean
ZIsMulti = ZIsMultiCell.Value
End Function

Private Function ZIsMultiCell() As Range
Dim R1&
R1 = ZR1
Set ZIsMultiCell = A_Cell.Worksheet.Cells(R1, ZIsMultiCno)
End Function

Private Function ZIsMultiCno%()
ZIsMultiCno = ZFndCno("IsMulti")
End Function

Private Function ZIsMultiValEntered() As Boolean
ZIsMultiValEntered = InStr(ZWrkCell_CurVal, vbLf) > 0
End Function

Private Function ZIsMust() As Boolean
ZIsMust = ZIsMustCell.Value
End Function

Private Function ZIsMustCell() As Range
Set ZIsMustCell = A_Cell.Worksheet.Cells(ZR1, ZIsMustCno)
End Function

Private Function ZIsMustCno%()
ZIsMustCno = ZFndCno("IsMust")
End Function

Private Function ZIsNoSelection() As Boolean
Dim R As Range
For Each R In ZAll_CharValName_Rge
    If R.Interior.Color = rgbYellow Then Exit Function
Next
ZIsNoSelection = True
End Function

Private Function ZIsSingle() As Boolean
ZIsSingle = Not ZIsMulti
End Function

Private Function ZIsVdt() As Boolean
If ZIsEr_Must Then Exit Function
If ZIsEr_MultiValEntered_ForSingle Then Exit Function
If ZIsEr_InvdtValEntered Then Exit Function
ZIsVdt = True
End Function

Private Function ZLeftCell_UnderSku() As Range
Dim C%
'C = ZLstCno_Sku
Set ZLeftCell_UnderSku = Ws_RC(A_Cell.Worksheet, A_Cell.Row, C)
End Function

Private Function ZMaxRno&()
Static X&
If X = 0 Then X = yWsMassUpd.Rows.Count
ZMaxRno = X
End Function

Private Function ZR1&()
If Not IsEmpty(ZLeftCell_UnderSku) Then ZR1 = A_Cell.Row: Exit Function
ZR1 = ZLeftCell_UnderSku.End(xlUp).Row
End Function

Private Function ZR2&()
Dim R&
R = ZLeftCell_UnderSku.End(xlDown).Row
If R = ZMaxRno Then
    ZR2 = ZRLast
Else
    ZR2 = R - 1
End If
End Function

Private Function ZRLast&()
ZRLast = Ws_RC(A_Cell.Worksheet, 1, ZCharValNameCno).End(xlDown).Row
End Function

Private Function ZVdtVal() As String()
Dim Sqv, O$(), J%, V$
Sqv = ZAll_CharValName_Rge.Value
For J = 1 To UBound(Sqv, 1)
    If Sqv(J, 1) <> "" Then Push O, Sqv(J, 1)
Next
ZVdtVal = O
End Function

Private Sub ZVdt_Again()
If ZIsVdt Then
    ZDoReset_TheCellColor
    ZDoReset_ErMsg
Else
    ZDoSet_TheCellToRed
    ZDoSet_ErMsg
End If
If ZIsEr_Must Then
    ZWrkCell.Value = "#MustInput#"
End If
End Sub

Private Function ZWrkAdrCell() As Range
Set ZWrkAdrCell = A_Cell.Worksheet.Cells(ZR1, ZWrkAdrCno)
End Function

Private Function ZWrkAdrCellVal$()
ZWrkAdrCellVal = ZWrkAdrCell.Value
End Function

Private Function ZWrkAdrCno%()
ZWrkAdrCno = ZFndCno("WrkAdr")
End Function

Private Function ZWrkCell() As Range
Set ZWrkCell = ZWrkWs.Range(ZWrkAdrCellVal)
End Function

Private Function ZWrkCell_CurVal$()
ZWrkCell_CurVal = ZWrkCell.Value
End Function

Private Function ZWrkCell_NewVal$()
ZWrkCell_NewVal = Join(ZAll_Selected_CharValName, vbLf)
End Function

Private Function ZWrkWs() As Worksheet
Set ZWrkWs = A_Cell.Worksheet.Parent.Sheets("Working")
End Function
