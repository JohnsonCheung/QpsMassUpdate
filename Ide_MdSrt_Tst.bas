Attribute VB_Name = "Ide_MdSrt_Tst"
Option Explicit
Dim A_Cell As Range

Sub ErWs_WsChg_V2DropDown(Cell As Range)
Set A_Cell = Cell
If ZIsCell_OutSideRange Then Exit Sub
If ZIsMulti Then
    ZCorrCell_ToggleVal_UsingSelVal
Else
    ZCorrCell_ReplVal_UsingSelVal
End If
ZDoSet_WrkWsCellVal
'------------------------------

ZVdt_Again           'Vdt again if any error.  Shw the message @ col-after-CharValName and first-row
End Sub

Private Property Get YAdr_LeftCell$()
YAdr_LeftCell$ = ZLeftCell_UnderSku.Address
End Property

Private Property Get YAdr_TheCell$()
YAdr_TheCell = ZWrkWsCell.Address
End Property

Private Property Get YAdr_WrkAdrCell$()
YAdr_WrkAdrCell = ZWrkAdrCell.Address
End Property

Private Property Get YCno_Sel%()
YCno_Sel = ZSelCno
End Property

Private Property Get YCno_WrkAdr%()
YCno_WrkAdr = ZWrkAdrCno
End Property

Private Property Get YIsMulti() As Boolean
YIsMulti = ZIsMulti
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

Private Function ZAll_CharValName_Rge() As Range
Set ZAll_CharValName_Rge = Ws_CRR(A_Cell.Worksheet, A_Cell.Column, ZR1, ZR2)
End Function

Private Function ZCorrCell() As Range
Set ZCorrCell = A_Cell.Cells(1, 0)
End Function

Private Sub ZCorrCell_Add_SelVal()
Dim A
    A = ZSelCell_Val
    If VarType(A) <> vbString Then Exit Sub
    If Trim(A) = "" Then Exit Sub

Dim B
    B = ZCorrCell_Val
    If VarType(A) <> vbString Then
        ZCorrCell_Val = A
        Exit Sub
    End If
    
Dim C$()
    C = Split(B, vbCrLf)
    Push C, A

ZCorrCell_Val = Join(C, vbCrLf)
End Sub

Private Function ZCorrCell_HasSelVal() As Boolean
Stop
ZCorrCell_HasSelVal = True
End Function

Private Sub ZCorrCell_ReplVal_UsingSelVal()
Stop
End Sub

Private Sub ZCorrCell_Rmv_SelVal()
Stop
ZDoUnlight_A_Cell A_Cell
End Sub

Private Sub ZCorrCell_ToggleVal_UsingSelVal()
If ZCorrCell_HasSelVal Then
    ZCorrCell_Rmv_SelVal
Else
    ZCorrCell_Add_SelVal
End If
ZCurRow_AdjHgt
End Sub

Private Property Get ZCorrCell_Val()
ZCorrCell_Val = ZCorrCell.Value
End Property

Private Property Let ZCorrCell_Val(V)

End Property

Private Sub ZCurRow_AdjHgt()

End Sub

Private Sub ZDoReset_ErMsg()
ZErMsgCell.Clear
End Sub

Private Sub ZDoReset_TheCellColor()
ZWrkWsCell.Font.ColorIndex = xlAutomatic
End Sub

Private Sub ZDoSet_ErMsg()
ZErMsgCell.Value = ZErMsg
End Sub

Private Sub ZDoSet_TheCellToRed()
ZWrkWsCell.Font.Color = rgbRed
End Sub

Private Sub ZDoSet_WrkWsCellVal()
With ZWrkWsCell
    .Interior.ColorIndex = xlNone
    .NumberFormat = "@"
    .Value = ZCorrCell_Val
End With
End Sub

Private Sub ZDoUnlight_A_Cell(Cell As Range)
Rge_RC(Cell, 1, 1).Interior.Pattern = xlNone
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
Ay = Split(ZWrkWsCell_CurVal, vbLf)
Vdt = ZVdtVal
For J = 0 To UB(Ay)
    If Ay_Idx(Vdt, Ay(J)) < 0 Then Push O, Ay(J)
Next
ZInvdtValEntered_StrAy = O
End Function

Private Function ZIsCell_NoValidation() As Boolean
ZIsCell_NoValidation = A_Cell.Validation.Type <> xlValidateList
End Function

Private Function ZIsCell_OutSideRange() As Boolean
ZIsCell_OutSideRange = True
If A_Cell.Row = 1 Then Exit Function
If A_Cell.Column <> ZSelCno Then Exit Function
If A_Cell.Row > ZRLast Then Exit Function
If ZIsCell_NoValidation Then Exit Function
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

Private Function ZIsMulti() As Boolean
ZIsMulti = ZMultiCell.Value = "Multi"
End Function

Private Function ZIsMultiValEntered() As Boolean
ZIsMultiValEntered = InStr(ZWrkWsCell_CurVal, vbLf) > 0
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
Stop
'C = ZLstCno_Sku
Set ZLeftCell_UnderSku = Ws_RC(A_Cell.Worksheet, A_Cell.Row, C)
End Function

Private Function ZMaxRno&()
Static X&
If X = 0 Then X = yWsMassUpd.Rows.Count
ZMaxRno = X
End Function

Private Function ZMultiCell() As Range
Set ZMultiCell = A_Cell.Worksheet.Cells(A_Cell.Row, ZMultiCno)
End Function

Private Function ZMultiCno%()
ZMultiCno = ZFndCno("Multi")
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
ZRLast = Ws_RC(A_Cell.Worksheet, 1, 1).End(xlDown).Row
End Function

Private Property Get ZSelCell_Val()
ZSelCell_Val = A_Cell.Value
End Property

Private Function ZSelCno%()
ZSelCno% = ZFndCno("Selection")
End Function

Private Function ZVdtVal() As String()
Dim A$
    A = A_Cell.Validation.Formula1
ZVdtVal = Split(A, vbCrLf)
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
    ZWrkWsCell.Value = "#MustInput#"
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

Private Function ZWrkWs() As Worksheet
Set ZWrkWs = A_Cell.Worksheet.Parent.Sheets("Working")
End Function

Private Function ZWrkWsCell() As Range
Set ZWrkWsCell = ZWrkWs.Range(ZWrkAdrCellVal)
End Function

Private Function ZWrkWsCell_CurVal$()
ZWrkWsCell_CurVal = ZWrkWsCell.Value
End Function
