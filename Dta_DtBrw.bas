Attribute VB_Name = "Dta_DtBrw"
Option Explicit
Enum eAlign
    eLeft = 1
    eCenter = 2
    eRight = 3
End Enum
Private A_Dt As TDt
Private ZFldNmAy$()
Private ZWdt%()
Private ZAlign() As eAlign
Private ZNFld%
Private ZNRow&
Private ZDrAy()

Sub DbT_Brw(T As DbT)
Dim O As TDt
    O.Nm = T.T
    O.FldNmAy = DbT_FldNmAy(T)
    O.DrAy = DbT_DrAy(T)
Dt_Brw O
End Sub

Sub DbT_Brw__Tst()
Dim D As Database
    Set D = OpenDatabase(Fs_Fb)
DbT_Brw DbT(D, "SkuCostChr")
D.Close
End Sub

Sub Dt_Brw(Dt As TDt, Optional NoRIdx As Boolean, Optional AlignAy)
A_Dt = Dt
ZDrAy = Dt.DrAy
ZFldNmAy = Dt.FldNmAy
ZNFld = Sz(ZFldNmAy)
ZNRow = Sz(ZDrAy)
ZWdt = ZZWdt

Dim J%
ReDim ZAlign(ZNFld - 1)
    If Not IsMissing(AlignAy) And Not IsEmpty(AlignAy) Then
        If Sz(AlignAy) <> ZNFld Then Er "Given {AlignAy-Len} should = {ZNFld}", Sz(AlignAy), ZNFld
        For J = 0 To UB(AlignAy)
            ZAlign(J) = AlignAy
        Next
    End If

Dim WithRIdx As Boolean
    WithRIdx = Not NoRIdx
    
Dim RIdxWdt%
    If WithRIdx Then
        RIdxWdt = Len(CStr(ZNRow))
    End If
    
Dim Oup_H1$
    Dim A_Ay$()
    Erase A_Ay
    Dim A_J%
    For A_J = 0 To ZNFld - 1
        Push A_Ay, String(ZWdt(A_J), "-")
    Next
    If WithRIdx Then
        Dim A1$, A2$
            A1 = String(RIdxWdt, "-")
            A2 = Join(A_Ay, " | ")
        Oup_H1 = Fmt_QQ("| ? | ? |", A1, A2)
    End If

Dim Oup_H2$
    Dim B_Ay$()
    Erase B_Ay
    For J = 0 To ZNFld - 1
        Push B_Ay, ZOup_Cell(ZFldNmAy, J)
    Next
    Dim B1$, B2$
        B1 = Space(RIdxWdt - 1) & "#"
        B2 = Join(B_Ay, " | ")
    Oup_H2 = Fmt_QQ("| ? | ? |", B1, B2)
    
Dim Oup_H3$
    Oup_H3 = Oup_H1

Dim Oup_Bottom$
    Oup_Bottom = Oup_H1

Dim O$()
    Push O, Oup_H1
    Push O, Oup_H2
    Push O, Oup_H3
    PushAy O, ZOup_Row(WithRIdx)
    Push O, Oup_Bottom

Str_Brw Join(O, vbCrLf)
End Sub

Sub Ws_DoBrwLstObj(Ws As Worksheet)
Dim LO As ListObject
    Set LO = Ws.ListObjects(1)
Dim Dt As TDt
    Dt.Nm = LO.Name
    Dt.FldNmAy = ZLO_FldNmAy(LO)
    Dt.DrAy = Sqv_DrAy(LO.DataBodyRange.Value)
Dt_Brw Dt
End Sub

Private Function DbT_DrAy(T As DbT)
DbT_DrAy = Rs_DrAy(T.D.TableDefs(T.T).OpenRecordset)
End Function

Private Function DbT_FldNmAy(T As DbT) As String()
DbT_FldNmAy = Flds_FldNmAy(T.D.TableDefs(T.T).Fields)
End Function

Private Function Flds_Dr(F As Fields) As Variant()
Dim U%
For U = F.Count - 1 To 0 Step -1
    If Not IsNull(F(U).Value) Then Exit For
Next
Dim J%
ReDim O(U)
For J = 0 To U
    O(J) = F(J).Value
Next
Flds_Dr = O
End Function

Private Function Flds_FldNmAy(F As Fields) As String()
Dim I As Field
Dim O$()
ReDim O(F.Count - 1)
Dim J%
J = 0
For Each I In F
    O(J) = I.Name
    J = J + 1
Next
Flds_FldNmAy = O
End Function

Private Function Rs_DrAy(Rs As Recordset)
Dim O()
    Dim NFld%
    Dim NRow&
    Dim R&, C%
    With Rs
        NFld = .Fields.Count
        NRow = .RecordCount
        While Not .EOF
            Push O, Flds_Dr(.Fields)
            .MoveNext
        Wend
    End With
Rs_DrAy = O
End Function

Private Function ZLO_FldNmAy(LO As ListObject) As String()
Dim V
    V = LO.HeaderRowRange.Value
Dim NFld%
    NFld = UBound(V, 2)
Dim O$()
    ReDim O(1 To NFld)
Dim J%
For J = 1 To NFld
    O(J) = V(1, J)
Next
ZLO_FldNmAy = O
End Function

Private Function ZOup_Align$(V, W%, Align As eAlign)
Dim S$
    S = V
Dim L%
    L = Len(S)
Dim O$
    If L > W Then
        Select Case W
        Case 1: O = "?"
        Case 2: O = "??"
        Case Else:  O = Left(S, W - 2) & ".."
        End Select
    ElseIf L = W Then
        O = S
    Else
        Dim A$
            A = Space(W - L)
        Select Case Align
        Case eRight: O = A & S
        Case eCenter:
            Dim A1$
                A1 = Space((W - L) \ 2)
            Dim A2$
                A2 = Space(W - L - Len(A1))
            O = A1 & S & A2
        Case Else: O = S & A
        End Select
    End If
ZOup_Align = O
End Function

Private Function ZOup_Cell$(Dr, C%)
Dim U%
    U = UB(Dr)
Dim V
    If 0 <= C And C <= U Then V = Dr(C)
Dim S$
    S = Str_Esc(V)
ZOup_Cell = ZOup_Align(S, ZWdt(C), ZAlign(C))
End Function

Private Function ZOup_OneRow$(R&)
Dim O$()
Dim I%
For I = 0 To ZNFld - 1
    Dim Dr
        Dr = ZDrAy(R)
    Push O, ZOup_Cell(Dr, I)
Next
ZOup_OneRow = "| " & Join(O, " | ") & " |"
End Function

Private Property Get ZOup_Row(WithRIdx As Boolean) As String()
Dim O$()
Dim J&
Dim W%
    W = Len(Str(ZNRow))
Dim A$
For J = 0 To ZNRow - 1
    If WithRIdx Then
        A = "| " & Space(W - Len(Str(J))) & J & " "
    End If
    Push O, A & ZOup_OneRow(J)
Next
ZOup_Row = O
End Property

Private Function ZZWdt() As Integer()
Dim J%
ReDim O%(ZNFld - 1)
For J = 0 To ZNFld - 1
    Dim W%
        W = Len(ZFldNmAy(J))
        Dim I&
        For I = 0 To ZNRow - 1
            Dim Dr
                Dr = ZDrAy(I)
            Dim U%
                U = Sz(Dr)
            Dim V
                V = Empty
                If 0 <= J And J <= U Then V = Dr(J)
            Dim L%
                L = Len(Str_Esc(V))
            If L > W Then W = L
        Next
    O(J) = W
Next
ZZWdt = O
End Function
