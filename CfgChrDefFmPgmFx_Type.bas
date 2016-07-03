Attribute VB_Name = "CfgChrDefFmPgmFx_Type"
Option Explicit

Type ChrDef     ' Defining a Char: What is its [Code Name IsMulti IsMust ValDic]
'    CtlType As eCtlTy
    CharCode As String
    CharName As String
    CostGp As String
    CostEle As String
    IsMulti As Boolean
    IsMust As Boolean
    IsNeedInList As Boolean ' That CtlType="Choose" or "Multivalue"
    Dic_OfValNm_ToValCd As Dictionary ' Key = ValName, Value=ValCode.  Using ValName as key is because, the Ws will use ValName to display and ValCode as "Val" is just for reference
End Type
