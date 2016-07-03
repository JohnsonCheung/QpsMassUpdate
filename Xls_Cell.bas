Attribute VB_Name = "Xls_Cell"
Option Explicit
Option Base 0

Sub Cell_AddCmt(Cell As Range, Cmt$, Optional Wdt% = 100, Optional Hgt% = 100)
If Cell_CmtTxt(Cell) <> "" Then Cell.Comment.Delete
Cell.AddComment Cmt
With Cell.Comment.Shape
    .Width = Wdt
    .Height = Hgt
End With
End Sub

Sub Cell_ClrBdr(Cell As Range)
With Cell
    .Borders(xlDiagonalDown).LineStyle = xlNone
    .Borders(xlDiagonalUp).LineStyle = xlNone
    .Borders(xlEdgeLeft).LineStyle = xlNone
    .Borders(xlEdgeTop).LineStyle = xlNone
    .Borders(xlEdgeBottom).LineStyle = xlNone
    .Borders(xlEdgeRight).LineStyle = xlNone
    .Borders(xlInsideVertical).LineStyle = xlNone
    .Borders(xlInsideHorizontal).LineStyle = xlNone
End With
End Sub

Function Cell_CmtTxt$(Cell As Range)
On Error Resume Next
Cell_CmtTxt = Rge_RC(Cell, 1, 1).Comment.Text
End Function

Sub Cell_Freeze(Cell As Range)
Dim mXls As Excel.Application: Set mXls = Cell.Application
Dim mWs As Worksheet: Set mWs = Cell.Worksheet:
Dim mWb As Workbook: Set mWb = mWs.Parent
'mXls.Goto Reference:="'" & mWs.Name & "'!" & Cell.Address(RowAbsolute:=True, COlumnAbsolute:=True, ReferenceStyle:=xlR1C1), Scroll:=True
mWb.Activate:
mWs.Select: mWs.Activate
Cell.Select: Cell.Activate
Dim mWin As Window: Set mWin = mXls.ActiveWindow
mWin.Activate
mWin.FreezePanes = True
End Sub

Sub Cell_Lnk(Cell As Range, CellTar As Range, Optional pScreenTip$ = "")
If pScreenTip = "" Then
    Cell.Hyperlinks.Add Cell, "", "'" & CellTar.Worksheet.Name & "'!" & CellTar.Address
Else
    Cell.Hyperlinks.Add Cell, "", "'" & CellTar.Worksheet.Name & "'!" & CellTar.Address, pScreenTip
End If
End Sub

Sub Cell_PutSqv(Cell As Range, Sqv)
If IsEmpty(Sqv) Then Exit Sub
Rge_RCRC(Cell, 1, 1, UBound(Sqv, 1), UBound(Sqv, 2)).Value = Sqv
End Sub
