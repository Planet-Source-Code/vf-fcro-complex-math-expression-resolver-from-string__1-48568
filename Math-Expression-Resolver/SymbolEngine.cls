VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SymbolEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private USERSYM() As String
Private USERSYMLEN() As Long 'BREAK SYMBOLS
Private USERCOUNT As Long

Private ENTERNODE() As String
Private ENTERNODELEN() As Long
Private EXITNODE() As String 'NODE SYMBOLS
Private EXITNODELEN() As Long
Private EXITENTERCOUNT As Long

Private NONSYM() As String
Private NONSYMLEN() As Long 'NON SYMBOLS
Private NONSYMCOUNT As Long


'SYMBOLIC NODE EXPRESSION ENGINE+++

'NON SYMBOLIC SINGLE CHARACTERS!
'SPACE
'TAB
'NEW LINE


'RADI Na PRINCIPU PRESIJECANJA (NODOVA), ODJELJIVANJA (BREAKS) i FALSE KARAKTERA KOJI SE NE PAMTE KAO REFERENCE

Event EnumSymbols(ByRef Symbol As Variant, ByRef ParentSymbol As Variant)
Event Finalize(ByRef SymbolReferences() As Variant, ByVal TotalCount As Long, ByRef ExpressionResult As Variant, ByRef ReportError As Boolean, ByRef ParentSymbol As Variant)


Public Property Let SetBreakSymbols(ByRef UserDefinedSymbols() As String)
Dim u As Long
USERSYM = UserDefinedSymbols
USERCOUNT = UBound(USERSYM)
ReDim USERSYMLEN(USERCOUNT)
For u = 0 To USERCOUNT
USERSYMLEN(u) = Len(USERSYM(u))
Next u
End Property

Public Function SetNodeSymbols(ByRef UserDefinedEnterNode() As String, ByRef UserDefinedExitNode() As String) As Boolean
On Error GoTo Dalje:
Dim u As Long
If UBound(UserDefinedEnterNode) <> UBound(UserDefinedExitNode) Then Exit Function
ENTERNODE = UserDefinedEnterNode
EXITNODE = UserDefinedExitNode
EXITENTERCOUNT = UBound(UserDefinedEnterNode)
ReDim ENTERNODELEN(EXITENTERCOUNT)
ReDim EXITNODELEN(EXITENTERCOUNT)
For u = 0 To EXITENTERCOUNT
ENTERNODELEN(u) = Len(ENTERNODE(u))
EXITNODELEN(u) = Len(EXITNODE(u))
Next u
SetNodeSymbols = True
Exit Function
Dalje:
On Error GoTo 0
End Function

Public Property Let SetNonSymbols(ByRef UserDefinedNonSymbols() As String)
Dim u As Long
NONSYM = UserDefinedNonSymbols
NONSYMCOUNT = UBound(NONSYM)
ReDim NONSYMLEN(NONSYMCOUNT)
For u = 0 To NONSYMCOUNT
NONSYMLEN(u) = Len(NONSYM(u))
Next u
End Property



Public Function ExpressionResolver(ByVal Expression As String, ByRef SymbolReaded As Long, ByRef IsError As Boolean, Optional ByVal ExitNodeIndex As Long = -1, Optional ByRef ParentSymbol As Variant = Empty) As Variant
On Error GoTo Dalje
Dim ExpressLen As Long
Dim RefCount As Long
Dim References() As Variant
Dim StringCount As Long
Dim SymbolStart As Long
Dim SymbolReads As Long
Dim ParSym As Variant
Dim u As Long

ReDim References(511)
ExpressLen = Len(Expression)


StringCount = 1
Kreni:
SymbolStart = StringCount
Slijedeci:

If ExitNodeIndex = -1 Then GoTo NemaNode

If CheckForSymbols(Expression, StringCount, EXITNODE(ExitNodeIndex), EXITNODELEN(ExitNodeIndex)) = 0 Then
    If SymbolStart = StringCount Then
    GoTo Enumeriraj
    Else
    References(RefCount) = Mid(Expression, SymbolStart, StringCount - SymbolStart)
    RaiseEvent EnumSymbols(References(RefCount), ParentSymbol)
    RefCount = RefCount + 1
    GoTo Enumeriraj
    End If
End If

NemaNode:

For u = 0 To USERCOUNT
'CHECK FOR SYMBOLS!
If CheckForSymbols(Expression, StringCount, USERSYM(u), USERSYMLEN(u)) = 0 Then

If SymbolStart = StringCount Then
StringCount = StringCount + USERSYMLEN(u)
References(RefCount) = USERSYM(u)
RefCount = RefCount + 1
GoTo Kreni

Else
References(RefCount) = Mid(Expression, SymbolStart, StringCount - SymbolStart)
RaiseEvent EnumSymbols(References(RefCount), ParentSymbol)
StringCount = StringCount + USERSYMLEN(u)
RefCount = RefCount + 1
References(RefCount) = USERSYM(u)
RefCount = RefCount + 1
GoTo Kreni

End If
End If
Next u



For u = 0 To EXITENTERCOUNT
If CheckForSymbols(Expression, StringCount, ENTERNODE(u), ENTERNODELEN(u)) = 0 Then
'CHECK ENTER NODE

If SymbolStart <> StringCount Then
References(RefCount) = Mid(Expression, SymbolStart, StringCount - SymbolStart)
RefCount = RefCount + 1
End If

StringCount = StringCount + ENTERNODELEN(u)

If RefCount <> 0 Then ParSym = References(RefCount - 1)
References(RefCount) = ExpressionResolver(Mid(Expression, StringCount), SymbolReads, IsError, u, ParSym)
If IsError Then Exit Function
If RefCount <> 0 Then References(RefCount - 1) = ParSym 'IF USER CHANGE PARENT SYMBOL THE NEW REFERENCE!
If VarType(References(RefCount)) <> 0 Then RefCount = RefCount + 1 'IF RETURN EMPTY REFERENCE SKIP THAT REFERENCE!
StringCount = StringCount + SymbolReads
GoTo Kreni
End If
Next u





'CHECK NON SYMBOLIC CHARACTERS [that means nothing,just does skip the character]
For u = 0 To NONSYMCOUNT
If CheckForSymbols(Expression, StringCount, NONSYM(u), NONSYMLEN(u)) = 0 Then
    If SymbolStart = StringCount Then
        StringCount = StringCount + NONSYMLEN(u)
    Else
        StringCount = StringCount + NONSYMLEN(u)
        References(RefCount) = Mid(Expression, SymbolStart, StringCount - SymbolStart - NONSYMLEN(u))
        RaiseEvent EnumSymbols(References(RefCount), ParentSymbol)
        RefCount = RefCount + 1
    End If
    GoTo Kreni
End If
Next u


If ExitNodeIndex <> -1 And StringCount >= ExpressLen Then IsError = True: Exit Function


If StringCount = ExpressLen Then
References(RefCount) = Mid(Expression, SymbolStart)
RaiseEvent EnumSymbols(References(RefCount), ParentSymbol)
RefCount = RefCount + 1
GoTo Enumeriraj
End If

StringCount = StringCount + 1

If StringCount > ExpressLen Then
    If ExitNodeIndex = -1 Then
    GoTo Enumeriraj
    Else
    IsError = True: Exit Function 'mora izaci sa EXIT NODE!
    End If
End If


GoTo Slijedeci

Enumeriraj:
IsError = False
RaiseEvent Finalize(References, RefCount, ExpressionResolver, IsError, ParentSymbol)
SymbolReaded = StringCount
Exit Function

Dalje:
On Error GoTo 0
IsError = True
End Function

'vrati 0=isti
Private Function CheckForSymbols(ByRef Expression As String, ByVal StartSearch As Long, ByRef Symbol As String, ByVal LenSymbol As Long) As Long
Dim TestP As String
TestP = Mid(Expression, StartSearch, LenSymbol)
CheckForSymbols = StrComp(TestP, Symbol, vbTextCompare)
End Function




