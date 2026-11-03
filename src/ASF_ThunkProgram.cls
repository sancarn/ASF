VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASF_ThunkProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Immutable after Build. Retains the AST so pointer-keyed cache entries cannot
' alias a later node. No scope, result, or invocation stack is cached here.
Private Type Instruction
    op As Long
    a As Long
    b As Long
    c As Long
End Type
Private code() As Instruction
Private quietErrors() As Boolean
Private quiet As Boolean
Private constants() As Variant
Private nodes As Collection
Private root As ASF_Map
Private used As Long, constantCount As Long

Friend Sub Build(ByVal node As ASF_Map, ByVal expression As Boolean)
    Set root = node
    Set nodes = New Collection
    ReDim code(0 To 63)
    ReDim quietErrors(0 To 63)
    ReDim constants(0 To 31)
    If expression Then
        EmitExpression node
    Else
        EmitStatement node, -1, -1
    End If
    Emit 0
End Sub

Friend Property Get Count() As Long
    Count = used
End Property
Friend Property Get PoolCount() As Long
    PoolCount = constantCount
End Property
Friend Property Get NodeAt(ByVal index As Long) As ASF_Map
    Set NodeAt = nodes(index)
End Property
Friend Property Get QuietFailure(ByVal pc As Long) As Boolean
    If pc >= 0 And pc < used Then QuietFailure = quietErrors(pc)
End Property
Friend Property Get IgnoreSignal(ByVal pc As Long) As Boolean
    IgnoreSignal = (code(pc).b = -2)
End Property
#If VBA7 And Win64 And Not Mac Then
Friend Property Get CodeAddress() As LongPtr
    CodeAddress = VarPtr(code(0))
End Property
Friend Property Get ConstantsAddress() As LongPtr
    ConstantsAddress = VarPtr(constants(0))
End Property
#End If

Private Function Emit(ByVal op As Long, Optional ByVal a As Long = 0, Optional ByVal b As Long = -1, Optional ByVal c As Long = -1) As Long
    If used > UBound(code) Then
        ReDim Preserve code(0 To used * 2 - 1)
        ReDim Preserve quietErrors(0 To used * 2 - 1)
    End If
    quietErrors(used) = quiet
    Emit = used
    With code(used)
        .op = op: .a = a: .b = b: .c = c
    End With
    used = used + 1
End Function
Private Function Keep(ByVal node As ASF_Map) As Long
    nodes.Add node
    Keep = nodes.Count
End Function
Private Sub Literal(ByVal value As Variant)
    If constantCount > UBound(constants) Then ReDim Preserve constants(0 To constantCount * 2 - 1)
    If IsObject(value) Then Set constants(constantCount) = value Else constants(constantCount) = value
    Emit 1, constantCount
    constantCount = constantCount + 1
End Sub

Private Sub EmitExpression(ByVal node As ASF_Map)
    If node Is Nothing Then Literal Empty: Exit Sub
    Dim kind As String: kind = node.GetValue("type")
    Dim op As String, branch As Long, finish As Long
    Select Case kind
        Case "Literal"
            Literal node.GetValue("value")
        Case "Binary"
            op = node.GetValue("op")
            EmitExpression node.GetValue("left")
            Emit 15 ' reference EvalBinary uses Let for the left operand
            Select Case op
                Case "&&", "||"
                    Emit 9
                    branch = Emit(IIf(op = "&&", 8, 16))
                    EmitExpression node.GetValue("right")
                    Emit 9
                    finish = Emit(7)
                    code(branch).a = used
                    Emit 13, IIf(op = "||", 1, 0)
                    code(finish).a = used
                Case "??"
                    branch = Emit(12)
                    EmitExpression node.GetValue("right")
                    code(branch).a = used
                Case Else
                    EmitExpression node.GetValue("right")
                    Emit 10, BinaryNumber(op)
            End Select
        Case "Unary"
            EmitExpression node.GetValue("expr")
            Emit 11, Keep(node)
        Case "Ternary"
            EmitExpression node.GetValue("cond")
            Emit 15 ' reference ternary condition uses Let
            Emit 9
            branch = Emit(8)
            EmitExpression node.GetValue("trueExpr")
            finish = Emit(7)
            code(branch).a = used
            EmitExpression node.GetValue("falseExpr")
            code(finish).a = used
        Case Else
            Emit 2, Keep(node)
    End Select
End Sub

Private Function BinaryNumber(ByVal op As String) As Long
    Select Case op
        Case "+": BinaryNumber = 0
        Case "-": BinaryNumber = 1
        Case "*": BinaryNumber = 2
        Case "/": BinaryNumber = 3
        Case "<": BinaryNumber = 4
        Case "<=": BinaryNumber = 5
        Case ">": BinaryNumber = 6
        Case ">=": BinaryNumber = 7
        Case "==", "=": BinaryNumber = 8
        Case "!=": BinaryNumber = 9
        Case "%": BinaryNumber = 10
        Case "^": BinaryNumber = 11
        Case "&": BinaryNumber = 12
        Case "|": BinaryNumber = 13
        Case "<<": BinaryNumber = 14
        Case ">>": BinaryNumber = 15
        Case Else: BinaryNumber = 16
    End Select
End Function

Private Sub EmitBlock(ByVal statements As Collection, ByVal breakTarget As Long, ByVal continueTarget As Long)
    Dim node As ASF_Map
    For Each node In statements
        EmitStatement node, breakTarget, continueTarget
    Next node
End Sub

Private Sub EmitStatement(ByVal node As ASF_Map, ByVal breakTarget As Long, ByVal continueTarget As Long)
    Dim kind As String: kind = node.GetValue("type")
    Dim branch As Long, finish As Long, loopStart As Long, loopBreak As Long, loopContinue As Long
    Dim conditions As Collection, blocks As Collection, jumps As Collection, item As Variant, i As Long
    Dim init As ASF_Map, stepNode As ASF_Map, cond As ASF_Map
    Select Case kind
        Case "Assign"
            EmitExpression node.GetValue("right")
            Emit 4, Keep(node.GetValue("left"))
        Case "ExprStmt"
            EmitExpression node.GetValue("expr")
            Emit 6
        Case "Return"
            EmitExpression node.GetValue("expr")
            Emit 5
        Case "Break", "Continue"
            i = IIf(kind = "Break", breakTarget, continueTarget)
            If i >= 0 Then Emit 7, i Else Emit 14, IIf(kind = "Break", 1, 2)
        Case "If"
            Set jumps = New Collection
            EmitExpression node.GetValue("cond")
            Emit 9
            branch = Emit(8)
            EmitBlock node.GetValue("then"), breakTarget, continueTarget
            jumps.Add Emit(7)
            code(branch).a = used
            Set conditions = node.GetValue("elseif_conds")
            Set blocks = node.GetValue("elseif_blocks")
            For i = 1 To conditions.Count
                EmitExpression conditions(i)
                Emit 9
                branch = Emit(8)
                EmitBlock blocks(i), breakTarget, continueTarget
                jumps.Add Emit(7)
                code(branch).a = used
            Next i
            If node.GetValue("hasElse") Then EmitBlock node.GetValue("else"), breakTarget, continueTarget
            For Each item In jumps
                code(CLng(item)).a = used
            Next item
        Case "While", "For"
            If kind = "For" Then
                If node.Exists("kind") Then
                    If LCase$(CStr(node.GetValue("kind"))) <> "classic" Then GoTo HostStatement
                End If
                Set init = node.GetValue("init")
                Set stepNode = node.GetValue("step")
                ' Reference classic-for discards init/step control signals, even ERR.
                If Not init Is Nothing Then Emit 3, Keep(AsStatement(init)), -2
            End If
            ' Forward exits use trampoline instructions, so host signals and
            ' nested break/continue share the same targets as explicit branches.
            branch = Emit(7)
            loopBreak = Emit(7)
            loopContinue = Emit(7)
            code(branch).a = used
            loopStart = used
            Set cond = node.GetValue("cond")
            If Not cond Is Nothing Then
                quiet = (kind = "For")
                EmitExpression cond
                Emit 9
                Emit 8, loopBreak
                quiet = False
            End If
            EmitBlock node.GetValue("body"), loopBreak, loopContinue
            code(loopContinue).a = used
            If Not stepNode Is Nothing Then Emit 3, Keep(AsStatement(stepNode)), -2
            Emit 7, loopStart
            code(loopBreak).a = used
        Case Else
HostStatement:
            Emit 3, Keep(node), breakTarget, continueTarget
    End Select
End Sub

Private Function AsStatement(ByVal node As ASF_Map) As ASF_Map
    Select Case CStr(node.GetValue("type"))
        Case "Assign", "Print", "If", "For", "While", "TryCatch", "Switch", "Return", "Break", "Continue", "ExprStmt"
            Set AsStatement = node
        Case Else
            Dim wrapper As New ASF_Map
            wrapper.SetValue "type", "ExprStmt"
            wrapper.SetValue "expr", node
            Set AsStatement = wrapper
    End Select
End Function
