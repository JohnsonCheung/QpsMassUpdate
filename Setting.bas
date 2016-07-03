Attribute VB_Name = "Setting"
Option Explicit

Const ZToGenCmpRpt_AppNm$ = "QpsMassUpdate"
Const ZToGenCmpRpt_Sec = "GenNewQQryFx"
Const ZToGenCmpRpt_Key_NewQDte$ = "NewQuoteDate"
Const ZToGenCmpRpt_Key_WrkFdr$ = "WrkFdr"
Const ZToGenCmpRpt_Key_MassUpdFxFn$ = "MassUpdFxFn"
Const ZToGenCmpRpt_Key_NewQQryFxFn$ = "NewQQryFxFn"

Sub GenCmpRpt_Get(NewQDte As Date, WrkFdr$, NewQQryFxFn$, MassUpdFxFn$)
NewQDte = GetSetting(ZToGenCmpRpt_AppNm, ZToGenCmpRpt_Sec, ZToGenCmpRpt_Key_NewQDte)
WrkFdr = GetSetting(ZToGenCmpRpt_AppNm, ZToGenCmpRpt_Sec, ZToGenCmpRpt_Key_WrkFdr)
NewQQryFxFn = GetSetting(ZToGenCmpRpt_AppNm, ZToGenCmpRpt_Sec, ZToGenCmpRpt_Key_NewQQryFxFn)
MassUpdFxFn = GetSetting(ZToGenCmpRpt_AppNm, ZToGenCmpRpt_Sec, ZToGenCmpRpt_Key_MassUpdFxFn)
End Sub

Sub GenCmpRpt_Sav(NewQDte As Date, WrkFdr$, NewQQryFxFn$, MassUpdFxFn$)
SaveSetting ZToGenCmpRpt_AppNm, ZToGenCmpRpt_Sec, ZToGenCmpRpt_Key_NewQDte, NewQDte
SaveSetting ZToGenCmpRpt_AppNm, ZToGenCmpRpt_Sec, ZToGenCmpRpt_Key_WrkFdr, WrkFdr
SaveSetting ZToGenCmpRpt_AppNm, ZToGenCmpRpt_Sec, ZToGenCmpRpt_Key_NewQQryFxFn, NewQQryFxFn
SaveSetting ZToGenCmpRpt_AppNm, ZToGenCmpRpt_Sec, ZToGenCmpRpt_Key_MassUpdFxFn, MassUpdFxFn
End Sub
