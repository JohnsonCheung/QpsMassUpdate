Attribute VB_Name = "SrcGenEnum"
Option Explicit
Enum eAA
    AA = 1&
    BB = 1
End Enum
Const LinFnEnd = "End Function"
Const CasEnd = "End Select"

Sub Src_GenEnm(Pj As VBProject, EnmLvs$)
Dim Ay$()
    Ay = SplitLvs(EnmLvs)
Dim Cmp As VBComponent
For Each Cmp In Pj.VBComponents
    Src_GenEnm__GenMd Cmp, Ay
Next
End Sub

Sub Src_GenEnm__Tst()
Src_GenEnm Application.VBE.ActiveVBProject, "ValTy ColTy FldTy WhichWs"
End Sub

Private Function Src_GenEnm__EnmAy(DclLin$) As Variant()
Dim O()
Dim Ay$()
Dim Lno%
Dim BegLno&
Dim EndLno&
Dim Lines$
Dim Fm%
Dim I%
Dim LinAy$()
Fm = 1
Ay = Split(DclLin, vbCrLf)
Do
    I = I + 1
    If I > 100 Then Stop
    BegLno = Src_GenEnm__LnoBegEnum(Ay, Fm)
    If BegLno = 0 Then Exit Do
    EndLno = Src_GenEnm__LnoEnd(Ay, BegLno + 1, "Enum")
    LinAy = Ay_BegEnd(Ay, BegLno - 1, EndLno - 1)
    Push O, LinAy
    Fm = EndLno + 1
Loop Until False
Src_GenEnm__EnmAy = O
End Function

Private Function Src_GenEnm__EnmNm$(Lin1$)
Dim S$
S = RmvPfx(Lin1, "Private ")
With Brk(S, " ")
    If .S1 <> "Enum" Then GoTo X
    If Left(.S2, 1) <> "e" Then GoTo X
    Src_GenEnm__EnmNm = Mid(.S2, 2)
End With
Exit Function
X:
Er "{Lin1} of given EnmLinAy must be [Enum e....]", Lin1
End Function

Private Function Src_GenEnm__FnLin$(EnmAy())
If Sz(EnmAy) = 0 Then Exit Function
Dim O$()
Dim EnmLinAy
For Each EnmLinAy In EnmAy
    Dim LinAy$()
        LinAy = EnmLinAy
    PushAy O, Src_GenEnm__OneFn(LinAy)
Next
Src_GenEnm__FnLin = Join(O, vbCrLf)
End Function

Private Sub Src_GenEnm__GenMd(Cmp As VBComponent, EnmNmAy$())
If Not (Cmp.Type = vbext_ct_StdModule Or Cmp.Type = vbext_ct_ClassModule) Then Exit Sub
Dim Md As CodeModule
Dim DclLin$
Dim EnmAy()
Dim FnLin$
Dim NDclLin%

Set Md = Cmp.CodeModule
NDclLin = Md.CountOfDeclarationLines:           If NDclLin = 0 Then Exit Sub
DclLin = Md.Lines(1, NDclLin)
EnmAy = Src_GenEnm__EnmAy(DclLin)
EnmAy = Src_GenEnm__SelEnmAy(EnmAy, EnmNmAy):   'If Sz(EnmAy) <> 0 Then Stop
FnLin = Src_GenEnm__FnLin(EnmAy):               'If FnLin <> "" Then Stop
Src_GenEnm__RmvEnmFn Md
If FnLin <> "" Then Md.InsertLines Md.CountOfLines + 1, FnLin
End Sub

Private Function Src_GenEnm__LnoBegEnum%(LinAy$(), FmLno%)
Dim O%
Dim L$

For O = FmLno To Sz(LinAy)
    L = RmvPfx(LinAy(O - 1), "Private ")
    If IsPfx(L, "Enum ") Then
        Src_GenEnm__LnoBegEnum = O
        Exit Function
    End If
Next
End Function

Private Function Src_GenEnm__LnoBegFn%(LinAy$())
Dim O%, L$
Dim I%
Dim Pfx$

For O = 0 To UB(LinAy)
    L = LinAy(O)
    For I = 1 To 2
        Select Case I
        Case 1: Pfx = "Function Enm_"
        Case 2: Pfx = "Private Function Enm_"
        End Select
        If IsPfx(L, Pfx) Then
            Src_GenEnm__LnoBegFn = O
            Exit Function
        End If
    Next
Nxt:
Next
End Function

Private Function Src_GenEnm__LnoEnd%(LinAy$(), Beg%, End_Function_or_Enum$)
Dim O%
For O = Beg To UB(LinAy)
    If LinAy(O) = "End " & End_Function_or_Enum Then Src_GenEnm__LnoEnd = O + 1: Exit Function
Next
Er "No [End Enum] line found in [LinAy] from {BegLno}", Beg
End Function

Private Function Src_GenEnm__OneFn(EnmLinAy$()) As String()
Dim EnmNm$
EnmNm$ = Src_GenEnm__EnmNm(EnmLinAy(0))
Dim U%
    U = UB(EnmLinAy) - 2
Dim MbrNmAy$()
Dim MbrValAy&()
    ReDim MbrNmAy(U)
    ReDim MbrValAy(U)
Dim J%
For J = 1 To UB(EnmLinAy) - 1
    With Brk(EnmLinAy(J), "=")
        MbrNmAy(J - 1) = .S1
        MbrValAy(J - 1) = .S2
    End With
Next

Dim UMbr%
    UMbr = UB(MbrNmAy)

Dim O$()
PushAy O, ZOupLinAy_FnToStr(EnmNm, UMbr, MbrNmAy, MbrValAy)
PushAy O, ZOupLinAy_FnFmStr(EnmNm, UMbr, MbrNmAy, MbrValAy)
Src_GenEnm__OneFn = O
End Function

Private Sub Src_GenEnm__OneFn__Tst()
Dim Inp$()
    Push Inp, "Private Enum eAA"
    Push Inp, "  AA = 1"
    Push Inp, "  BB = 2"
    Push Inp, "End Enum"
Ay_Brw Src_GenEnm__OneFn(Inp)
End Sub

Private Sub Src_GenEnm__RmvEnmFn(Md As CodeModule)
Dim LinAy$()
Dim I%
Dim LnoBeg%, LnoEnd%
Dim HasRmv As Boolean
Do
    I = I + 1
    If I = 100 Then Stop
    LinAy = Split(Md.Lines(1, Md.CountOfLines), vbCrLf)
    LnoBeg = Src_GenEnm__LnoBegFn(LinAy)
    If LnoBeg = 0 Then
        If HasRmv Then
            Debug.Print Md.Parent.Name
        End If
        Exit Sub
    End If
    LnoEnd = Src_GenEnm__LnoEnd(LinAy, LnoBeg + 1, "Function")
    Md.DeleteLines LnoBeg, LnoEnd - LnoBeg + 1
    HasRmv = True
Loop Until False
End Sub

Private Function Src_GenEnm__SelEnmAy(EnmAy(), EnmNmAy$()) As Variant()
'Select element in EnmAy having names as in EnmNmAy
'Each element in EnmAy is EnmLinAy$()
Dim O()
Dim J%
For J = 0 To UB(EnmAy)
    Dim EnmNm$
        EnmNm = Src_GenEnm__EnmNm(CStr(EnmAy(J)(0)))
    If Ay_Has(EnmNmAy, EnmNm) Then
        Push O, EnmAy(J)
    End If
Next
Src_GenEnm__SelEnmAy = O
End Function

Private Function ZOupLinAy(LinFn$, LinDimO$, CasSel$, CasLinAy$(), CasElse$, LinSetRet$) As String()
Dim O$()
  Push O, LinFn
  Push O, LinDimO
  Push O, CasSel
PushAy O, CasLinAy
  Push O, CasElse
  Push O, CasEnd
  Push O, LinSetRet
  Push O, LinFnEnd
  Push O, ""
ZOupLinAy = O
End Function

Private Property Get ZOupLinAy_FnFmStr(EnmNm$, UMbr%, MbrNmAy$(), MbrValAy&()) As String()
Dim FnNm$
Dim J%, Spc_S$, Spc_Max%, Spc_Cur%
Dim CasLinAy$()
Dim LinFn$
Dim LinDimO$
Dim CasSel$
Dim CasElse$
Dim LinSetRet$
Dim A$, B$, C$

FnNm = Fmt_QQ("Enm_?", EnmNm)
ReDim CasLinAy(UMbr)
Spc_Max = Ay_MaxLen(MbrNmAy)
For J = 0 To UMbr
    Spc_Cur = Len(MbrNmAy(J))
    Spc_S = Space(Spc_Max - Spc_Cur)
    CasLinAy(J) = Fmt_QQ("Case ""?"": ?O = e?.?", MbrNmAy(J), Spc_S, EnmNm, MbrNmAy(J))
Next
LinFn = Fmt_QQ("Function ?(S$) as e?", FnNm, EnmNm)
LinDimO = Fmt_QQ("Dim O As e?", EnmNm)
CasSel = "Select Case S"
A = EnmNm
B = Quote(Join(MbrNmAy), """[*]""")
C = "Case Else: Er ""Given {S} is a not in valid Enm-e?-{MbrNmList}"",S,?"
CasElse = Fmt_QQ(C, A, B)
LinSetRet = Fmt_QQ("? = O", FnNm)
ZOupLinAy_FnFmStr = ZOupLinAy(LinFn, LinDimO, CasSel, CasLinAy, CasElse, LinSetRet)
End Property

Private Property Get ZOupLinAy_FnToStr(EnmNm$, UMbr%, MbrNmAy$(), MbrValAy&()) As String()
Dim FnNm$
Dim CasLinAy$()
Dim J%
Dim Spc_S$, Spc_Max%, Spc_Cur%
Dim LinFn$
Dim LinDimO$
Dim CasSel$
Dim CasElse$
Dim MbrValList$
Dim MbrNmList$

FnNm = Fmt_QQ("Enm_?_ToStr", EnmNm)
ReDim CasLinAy(UMbr)
Spc_Max = Ay_MaxLen(MbrNmAy)
For J = 0 To UMbr
    Spc_Cur = Len(MbrNmAy(J))
    Spc_S = Space(Spc_Max - Spc_Cur)
    CasLinAy(J) = Fmt_QQ("Case e?.?: ?O = ""?""", EnmNm, MbrNmAy(J), Spc_S, MbrNmAy(J))
Next
LinFn = Fmt_QQ("Function ?(P As e?)", FnNm, EnmNm)
LinDimO = "Dim O$"
CasSel = "Select Case P"
MbrValList = """" & Quote(Ay_Join(MbrValAy), "[]") & """"
MbrNmList = """" & Quote(Join(MbrNmAy), "[]") & """"
Const C = "Case Else: Er ""Enm-e?-{MbrVal} not in valid {MbrVal-List} of {MbrNm-List}"", P,?,?"
CasElse = Fmt_QQ(C, EnmNm, MbrValList, MbrNmList)
Dim LinSetRet$: LinSetRet = Fmt_QQ("? = O", FnNm)
ZOupLinAy_FnToStr = ZOupLinAy(LinFn, LinDimO, CasSel, CasLinAy, CasElse, LinSetRet)
End Property
