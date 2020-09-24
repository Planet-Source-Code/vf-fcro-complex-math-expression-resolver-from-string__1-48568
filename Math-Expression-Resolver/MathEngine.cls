VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MathPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim SYMBOLS() As String
Private WithEvents SY As SymbolEngine
Attribute SY.VB_VarHelpID = -1

Public Property Get MathExpression(ByRef Expression As String, ByRef IsError As Boolean) As Variant
Dim Readed As Long
MathExpression = SY.ExpressionResolver(Expression, Readed, IsError)
End Property




Private Function MathEngine(ByRef ExpressionReferences() As Variant, ByVal TotalCount As Long, ByRef IsError As Boolean) As Variant
On Error GoTo Dalje
Dim VarT As Long
Dim u As Long
Dim TempsCounter As Long
Dim LastPos As Long
Dim TempVal As Double
TempVal = 1

'ORIGANIZE POSITIVE AND NEGATIVE NUMBERS FROM SYMBOLS
LastPos = TotalCount - 1
For u = 0 To LastPos

If u <> LastPos Then
'CHECK IF NUMBER FOLLOW NUMBER-!...Error
If IsNumeric(ExpressionReferences(u)) And IsNumeric(ExpressionReferences(u + 1)) Then IsError = True: Exit Function
End If


If StrComp(ExpressionReferences(u), SYMBOLS(0), vbTextCompare) = 0 Then

If u = LastPos Then IsError = True: Exit Function

    
ElseIf StrComp(ExpressionReferences(u), SYMBOLS(1), vbTextCompare) = 0 Then

    If u = LastPos Then IsError = True: Exit Function
    TempVal = TempVal * -1
    
ElseIf StrComp(ExpressionReferences(u), SYMBOLS(2), vbTextCompare) = 0 Then
GoTo InRestC
 
ElseIf StrComp(ExpressionReferences(u), SYMBOLS(3), vbTextCompare) = 0 Then
GoTo InRestC
 
ElseIf StrComp(ExpressionReferences(u), SYMBOLS(4), vbTextCompare) = 0 Then
GoTo InRestC
 
ElseIf StrComp(ExpressionReferences(u), SYMBOLS(5), vbTextCompare) = 0 Then
GoTo InRestC

ElseIf StrComp(ExpressionReferences(u), SYMBOLS(6), vbTextCompare) = 0 Then
GoTo InRestC

ElseIf StrComp(ExpressionReferences(u), SYMBOLS(7), vbTextCompare) = 0 Then
GoTo InRestC

ElseIf StrComp(ExpressionReferences(u), SYMBOLS(18), vbTextCompare) = 0 Then
GoTo InRestC

ElseIf StrComp(ExpressionReferences(u), SYMBOLS(19), vbTextCompare) = 0 Then
GoTo InRestC

ElseIf StrComp(ExpressionReferences(u), SYMBOLS(8), vbTextCompare) = 0 Then
InRestC:
    If u = LastPos Or u = 0 Then IsError = True: Exit Function
    If TempsCounter <> u Then ExpressionReferences(TempsCounter) = ExpressionReferences(u)
    TempVal = 1
    TempsCounter = TempsCounter + 1
    
Else

    VarT = VarType(ExpressionReferences(u))
    If VarT = 8 Then
    'IF IS STRING = CHECK IF IS NUMBER?!
        If IsNumeric(ExpressionReferences(u)) Then
        ExpressionReferences(TempsCounter) = TempVal * Val(ExpressionReferences(u))
        Else
        IsError = True: Exit Function
        End If
       
    ElseIf VarT = 5 Then
    'IF IS DOUBLE
    ExpressionReferences(TempsCounter) = TempVal * ExpressionReferences(u)
    End If

    TempVal = 1
    TempsCounter = TempsCounter + 1
    
End If

Next u


'EXECUTE MTP
LastPos = TempsCounter - 1
TempsCounter = 0

For u = 0 To LastPos
VarT = VarType(ExpressionReferences(u))
If VarT = 8 Then
    If StrComp(ExpressionReferences(u), SYMBOLS(4), vbTextCompare) = 0 Then
    ExpressionReferences(TempsCounter - 1) = ExpressionReferences(TempsCounter - 1) ^ ExpressionReferences(u + 1)
    u = u + 1
    Else
    If TempsCounter <> u Then ExpressionReferences(TempsCounter) = ExpressionReferences(u)
    TempsCounter = TempsCounter + 1
    End If
Else
If TempsCounter <> u Then ExpressionReferences(TempsCounter) = ExpressionReferences(u)
TempsCounter = TempsCounter + 1
End If
Next u



'EXECUTE MULTIPLY AND DIVIDE
LastPos = TempsCounter - 1
TempsCounter = 0
For u = 0 To LastPos
VarT = VarType(ExpressionReferences(u))
If VarT = 8 Then
    If StrComp(ExpressionReferences(u), SYMBOLS(3), vbTextCompare) = 0 Then
    ExpressionReferences(TempsCounter - 1) = ExpressionReferences(TempsCounter - 1) / ExpressionReferences(u + 1)
    u = u + 1
    ElseIf StrComp(ExpressionReferences(u), SYMBOLS(2), vbTextCompare) = 0 Then
    ExpressionReferences(TempsCounter - 1) = ExpressionReferences(TempsCounter - 1) * ExpressionReferences(u + 1)
    u = u + 1
    Else
    If TempsCounter <> u Then ExpressionReferences(TempsCounter) = ExpressionReferences(u)
    TempsCounter = TempsCounter + 1
    End If
Else
If TempsCounter <> u Then ExpressionReferences(TempsCounter) = ExpressionReferences(u)
TempsCounter = TempsCounter + 1
End If
Next u


'ZBROJI PRIJE LOGICKIH
LastPos = TempsCounter - 1
TempVal = 0
If LastPos = 0 Then GoTo AvoidAdd

TempsCounter = 0
For u = 0 To LastPos

If VarType(ExpressionReferences(u)) = 5 Then
    TempVal = TempVal + ExpressionReferences(u)
    
Else

    ExpressionReferences(TempsCounter) = TempVal
    TempsCounter = TempsCounter + 1
    ExpressionReferences(TempsCounter) = ExpressionReferences(u)
    TempsCounter = TempsCounter + 1
    TempVal = 0
    
End If

If u = LastPos Then
ExpressionReferences(TempsCounter) = TempVal
TempsCounter = TempsCounter + 1
End If

Next u

AvoidAdd:



'TREBA IZVRŠITI ARITMETIÈKE OPERACIJE PRIJE LOGICKIH
LastPos = TempsCounter - 1
TempsCounter = 0
For u = 0 To LastPos
VarT = VarType(ExpressionReferences(u))
If VarT = 8 Then
    If StrComp(ExpressionReferences(u), SYMBOLS(5), vbTextCompare) = 0 Then
    ExpressionReferences(TempsCounter - 1) = ExpressionReferences(TempsCounter - 1) And ExpressionReferences(u + 1)
    u = u + 1
    ElseIf StrComp(ExpressionReferences(u), SYMBOLS(6), vbTextCompare) = 0 Then
    ExpressionReferences(TempsCounter - 1) = ExpressionReferences(TempsCounter - 1) Or ExpressionReferences(u + 1)
    u = u + 1
    ElseIf StrComp(ExpressionReferences(u), SYMBOLS(7), vbTextCompare) = 0 Then
    ExpressionReferences(TempsCounter - 1) = ExpressionReferences(TempsCounter - 1) Mod ExpressionReferences(u + 1)
    u = u + 1
    ElseIf StrComp(ExpressionReferences(u), SYMBOLS(8), vbTextCompare) = 0 Then
    ExpressionReferences(TempsCounter - 1) = ExpressionReferences(TempsCounter - 1) Xor ExpressionReferences(u + 1)
    u = u + 1
    ElseIf StrComp(ExpressionReferences(u), SYMBOLS(18), vbTextCompare) = 0 Then
    ExpressionReferences(TempsCounter - 1) = ExpressionReferences(TempsCounter - 1) Eqv ExpressionReferences(u + 1)
    u = u + 1
    ElseIf StrComp(ExpressionReferences(u), SYMBOLS(19), vbTextCompare) = 0 Then
    ExpressionReferences(TempsCounter - 1) = ExpressionReferences(TempsCounter - 1) Imp ExpressionReferences(u + 1)
    u = u + 1
    
    Else
    If TempsCounter <> u Then ExpressionReferences(TempsCounter) = ExpressionReferences(u)
    TempsCounter = TempsCounter + 1
    End If
Else
If TempsCounter <> u Then ExpressionReferences(TempsCounter) = ExpressionReferences(u)
TempsCounter = TempsCounter + 1
End If
Next u


'ZADNJE ZBRAJANJE PREOSTALIH BROJEVA
TempVal = 0
For u = 0 To TempsCounter - 1
TempVal = TempVal + ExpressionReferences(u)
Next u

MathEngine = TempVal
IsError = False
Exit Function

Dalje:
On Error GoTo 0
IsError = True
End Function

Private Sub Class_Initialize()
Dim u As Long
Dim NodeENTERS() As String
Dim NodeEXITS() As String
Dim NoSYM() As String
Set SY = New SymbolEngine

ReDim SYMBOLS(19) 'BREAK SYMBOLS
SYMBOLS(0) = "+"
SYMBOLS(1) = "-"
SYMBOLS(2) = "*"
SYMBOLS(3) = "/"
SYMBOLS(4) = "^"

SYMBOLS(5) = "AND"
SYMBOLS(6) = "OR"
SYMBOLS(7) = "MOD"
SYMBOLS(8) = "XOR"

SYMBOLS(9) = "SQR"
SYMBOLS(10) = "TAN"
SYMBOLS(11) = "ATN"
SYMBOLS(12) = "SIN"
SYMBOLS(13) = "COS"
SYMBOLS(14) = "EXP"
SYMBOLS(15) = "LOG"
SYMBOLS(16) = "ABS"
SYMBOLS(17) = "SGN"

SYMBOLS(18) = "EQV"
SYMBOLS(19) = "IMP"


ReDim NodeENTERS(0)
ReDim NodeEXITS(0)

NodeENTERS(0) = "(" 'NODE SYMBOLS
NodeEXITS(0) = ")"

ReDim NoSYM(2)
NoSYM(0) = vbTab
NoSYM(1) = vbCrLf
NoSYM(2) = " "
SY.SetNonSymbols = NoSYM
SY.SetBreakSymbols = SYMBOLS
SY.SetNodeSymbols NodeENTERS, NodeEXITS

End Sub

Private Sub Class_Terminate()
Set SY = Nothing
End Sub



Private Sub SY_Finalize(SymbolReferences() As Variant, ByVal TotalCount As Long, ExpressionResult As Variant, ReportError As Boolean, ParentSymbol As Variant)
On Error GoTo Dalje
If StrComp(ParentSymbol, SYMBOLS(9), vbTextCompare) = 0 Then

'vrati umjesto SQR reference rezultat,a kao referencu unutar NODA () vrati EMPTY da ga preskoci!
'Ovim se može objediniti sintaksa naredbe:
'NEKISYM (....) koji se sastoji od 2 reference: NEKISYM i (....) onoga u zagradi kao rezultata.."
'Ako vratimo kao rezultat NEKISYM ostaje kao referenca,a (....) daje rezultat!!!!


ParentSymbol = Sqr(MathEngine(SymbolReferences, TotalCount, ReportError))

ElseIf StrComp(ParentSymbol, SYMBOLS(10), vbTextCompare) = 0 Then
ParentSymbol = Tan(MathEngine(SymbolReferences, TotalCount, ReportError))
ElseIf StrComp(ParentSymbol, SYMBOLS(11), vbTextCompare) = 0 Then
ParentSymbol = Atn(MathEngine(SymbolReferences, TotalCount, ReportError))

ElseIf StrComp(ParentSymbol, SYMBOLS(12), vbTextCompare) = 0 Then
ParentSymbol = Sin(MathEngine(SymbolReferences, TotalCount, ReportError))
ElseIf StrComp(ParentSymbol, SYMBOLS(13), vbTextCompare) = 0 Then
ParentSymbol = Cos(MathEngine(SymbolReferences, TotalCount, ReportError))

ElseIf StrComp(ParentSymbol, SYMBOLS(14), vbTextCompare) = 0 Then
ParentSymbol = Exp(MathEngine(SymbolReferences, TotalCount, ReportError))
ElseIf StrComp(ParentSymbol, SYMBOLS(15), vbTextCompare) = 0 Then
ParentSymbol = Log(MathEngine(SymbolReferences, TotalCount, ReportError))

ElseIf StrComp(ParentSymbol, SYMBOLS(16), vbTextCompare) = 0 Then
ParentSymbol = Val(Abs(MathEngine(SymbolReferences, TotalCount, ReportError)))
ElseIf StrComp(ParentSymbol, SYMBOLS(17), vbTextCompare) = 0 Then
ParentSymbol = Val(Sgn(MathEngine(SymbolReferences, TotalCount, ReportError)))

ElseIf IsNumeric(ParentSymbol) And ParentSymbol <> Empty Then
ReportError = True

Else
ExpressionResult = MathEngine(SymbolReferences, TotalCount, ReportError)

End If
Exit Sub
Dalje:
On Error GoTo 0
ReportError = True
End Sub
