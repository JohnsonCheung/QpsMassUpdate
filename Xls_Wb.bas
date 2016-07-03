Attribute VB_Name = "Xls_Wb"
Option Explicit

Function Wb_AddName(Wb As Workbook, Nm$, Rge As Range) As Name
Dim WsNm$
    WsNm = Rge_Ws(Rge).Name
Dim Formula$
    Formula = Fmt_QQ("='?'!?", WsNm, Rge.Address)
Set Wb_AddName = Wb.Names.Add(Nm, Formula)
End Function

Function Wb_AddWs_AtEnd(Wb As Workbook, WsNm$, Optional DltBefAdd As Boolean) As Worksheet
If DltBefAdd Then Wb_DltWs Wb, WsNm
Dim O As Worksheet
Set O = Wb.Worksheets.Add(, Wb.Sheets(Wb.Sheets.Count))
O.Name = WsNm
Debug.Print Wb.Application.VBE.ActiveVBProject.VBComponents.Count
If O.CodeName = "" Then
    Dim A%
    A = Wb.Application.VBE.ActiveVBProject.VBComponents.Count
    If O.CodeName = "" Then Stop
End If
Set Wb_AddWs_AtEnd = O
End Function

Sub Wb_AssertWsExist(Wb As Workbook, WsNm$)
If Wb_IsWs(Wb, WsNm) Then Exit Sub
Const C = "{Worksheet} not found in {Workbook} in {folder}"
Er C, WsNm, Wb.Name, Wb_Pth(Wb)
End Sub

Sub Wb_Chk_IsWs(Wb As Workbook, WsNm$, OEr$())
If Wb_IsWs(Wb, WsNm) Then Exit Sub
Push OEr, Fmt_QQ("Wb(?) does not have Ws(?)", Wb.Name, WsNm)
End Sub

Sub Wb_ClrNames(Wb As Workbook, pPfx$)
Dim J%
Dim L%: L = Len(pPfx)
For J = Wb.Names.Count To 1 Step -1
    Dim iNm As Name: Set iNm = Wb.Names(J)
    Dim mNm$: mNm = iNm.Name
    Dim mNmX$
    Dim mP%: mP = InStr(mNm, "!")
    If mP > 0 Then
        mNmX = Mid(mNm, mP + 1)
    Else
        mNmX = mNm
    End If
    If pPfx = "" Then
        iNm.Delete
    Else
        If Left(mNmX, L) = pPfx Then iNm.Delete
    End If
Next
End Sub

Sub Wb_DltWs(Wb As Workbook, WsNm$)
Dim J%
For J% = 1 To Wb.Sheets.Count
    Dim IWs As Worksheet: Set IWs = Wb.Sheets(J)
    If IWs.Name = WsNm Then Ws_Dlt IWs: Exit Sub
Next
End Sub

Sub Wb_Hid_WsNmAy(Wb As Workbook, WsNmAy$())
Dim Ws As Worksheet, J%
For J = 0 To UB(WsNmAy)
    If Wb_IsWs(Wb, WsNmAy(J)) Then
        Set Ws = Wb.Sheets(WsNmAy(J))
        Ws.Visible = xlSheetHidden
    End If
Next
End Sub

Function Wb_IsWs(Wb As Workbook, WsNm$) As Boolean
Dim Ws As Worksheet
For Each Ws In Wb.Sheets
    If Ws.Name = WsNm Then Wb_IsWs = True: Exit Function
Next
End Function

Function Wb_Lik(LikStr$, Optional OCnt%) As Workbook
OCnt = 0
Dim Wb As Workbook
For Each Wb In Workbooks
    If Wb.Name Like LikStr Then
        OCnt = OCnt + 1
        Set Wb_Lik = Wb
    End If
Next
End Function

Function Wb_New() As Workbook
Set Wb_New = Workbooks.Add
End Function

Function Wb_Pth$(Wb As Workbook)
Wb_Pth = Ffn_Pth(Wb.FullName)
End Function

Function Wb_Ws(Wb As Workbook, WsNm_or_Idx) As Worksheet
Set Wb_Ws = Wb.Sheets(WsNm_or_Idx)
End Function

Function Wb_WsOpt(Wb As Workbook, WsNm_or_Idx) As Worksheet
On Error Resume Next
Set Wb_WsOpt = Wb.Sheets(WsNm_or_Idx)
End Function
