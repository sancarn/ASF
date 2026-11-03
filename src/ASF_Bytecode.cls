VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ASF_Bytecode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Compiler for the Garcia architecture experiment. Names are interned once.
' Dynamic binding IDs preserve ASF's nearest-existing-caller-scope assignment.
Private op() As Long, a() As Long, b() As Long, handlers() As Long
Private constants() As Variant, names() As String
Private entries() As Long, parameters() As Variant
Private used As Long, poolUsed As Long, nameUsed As Long
Private globals As ASF_Globals, pending As Collection, functions As Collection
Private currentHandler As Long, folded As Long, fused As Long
Private rejection As String, optimise As Boolean
Private scratchCount As Long
Private scoped As Boolean

Friend Function Build(ByVal state As ASF_Globals, ByVal index As Long, ByVal optimize As Boolean) As Boolean
    On Error GoTo Unsupported
    Set globals = state: optimise = optimize
    Set pending = New Collection: Set functions = New Collection
    ReDim op(0 To 127): ReDim a(0 To 127): ReDim b(0 To 127): ReDim handlers(0 To 127)
    ReDim constants(0 To 31): ReDim names(0 To 31)
    ReDim entries(0 To state.gPrograms.Count): ReDim parameters(0 To state.gPrograms.Count)
    currentHandler = -1
    Dim program As Variant, scope As Collection, frame As ASF_Map, stmt As ASF_Map
    program = state.gPrograms(index)
    Set scope = program(2)
    For Each frame In scope
        If frame.Count <> 0 Then Reject "nonempty raw scope template"
    Next frame
    CompileBlock program(1), -1, -1
    Emit 0 ' falling off main leaves OUTPUT_ unchanged
    Dim nextFunction As Long, fn As Long, args As Variant, ids() As Long, i As Long, nm As String
    nextFunction = 1
    Do While nextFunction <= pending.Count
        fn = pending(nextFunction): nextFunction = nextFunction + 1
        program = globals.gPrograms(fn): nm = CStr(program(0))
        entries(fn) = used
        ReDim ids(0 To 2): ids(0) = 0: ids(1) = -1
        If globals.gFuncParams.TryGetValue(nm, args) Then
            If Not IsArray(args) Then Reject "non-array named-function parameters"
            ReDim ids(0 To UBound(args) - LBound(args) + 3)
            ids(0) = UBound(ids) - 2
            ids(ids(0) + 1) = -1
            For i = 1 To ids(0)
                ids(i) = NameID(CStr(args(LBound(args) + i - 1)))
            Next i
        End If
        ids(ids(0) + 2) = ids(0)
        If ids(0) = 1 Then
            If names(ids(1)) = vbNullString Then ids(ids(0) + 2) = 0
        End If
        If UBound(program) >= 3 Then
            If VarType(program(3)) = vbString Then
                If Len(program(3)) > 0 Then ids(ids(0) + 1) = NameID(CStr(program(3)))
            End If
        End If
        parameters(fn) = ids
        CompileBlock program(1), -1, -1
        Literal Empty: Emit 6 ' implicit function return
    Loop
    ' Shadowed named functions need the reference's dynamic callable resolution.
    For i = 0 To nameUsed - 1
        If globals.gFuncTable.Exists(names(i)) Then Reject "named function used as a value or shadowed: " & names(i)
    Next i
    ReDim Preserve op(0 To used - 1): ReDim Preserve a(0 To used - 1)
    ReDim Preserve b(0 To used - 1): ReDim Preserve handlers(0 To used - 1)
    ReDim Preserve names(0 To IIf(nameUsed = 0, 0, nameUsed - 1))
    Build = True
    Exit Function
Unsupported:
    If Len(rejection) = 0 Then rejection = "lowering error: " & Err.Description
    Err.Clear
End Function
Friend Property Get Reason() As String
    Reason = rejection
End Property
Friend Function Export() As Variant
    Export = Array(op, a, b, handlers, constants, names, entries, parameters, nameUsed, folded, fused)
End Function
Private Sub Reject(ByVal reason As String)
    rejection = reason
    Err.Raise 5, "ASF.Bytecode", reason
End Sub
Private Function Emit(ByVal instruction As Long, Optional ByVal operand As Long = 0, Optional ByVal target As Long = 0) As Long
    If used > UBound(op) Then
        ReDim Preserve op(0 To used * 2 - 1): ReDim Preserve a(0 To used * 2 - 1)
        ReDim Preserve b(0 To used * 2 - 1): ReDim Preserve handlers(0 To used * 2 - 1)
    End If
    Emit = used: op(used) = instruction: a(used) = operand: b(used) = target
    handlers(used) = currentHandler: used = used + 1
End Function
Private Function Pool(ByVal value As Variant) As Long
    If poolUsed > UBound(constants) Then ReDim Preserve constants(0 To poolUsed * 2 - 1)
    If IsObject(value) Then Set constants(poolUsed) = value Else constants(poolUsed) = value
    Pool = poolUsed: poolUsed = poolUsed + 1
End Function
Private Sub Literal(ByVal value As Variant)
    Emit 1, Pool(value)
End Sub
Private Function NameID(ByVal name As String) As Long
    Dim i As Long
    For i = 0 To nameUsed - 1
        If names(i) = name Then NameID = i: Exit Function
    Next i
    If Not scoped Then
    If name = "Math" Or name = "this" Or name = "super" Then Reject "special binding: " & name
    If Left$(name, 2) = "__" Then Reject "reserved runtime binding: " & name
    If InStr(name, ".") > 0 Or InStr(name, "[") > 0 Then Reject "legacy path binding: " & name
    End If
    If nameUsed > UBound(names) Then ReDim Preserve names(0 To nameUsed * 2 - 1)
    names(nameUsed) = name: NameID = nameUsed: nameUsed = nameUsed + 1
End Function
Private Function BinaryID(ByVal symbol As String) As Long
    Select Case symbol
        Case "+": BinaryID = 20
        Case "-": BinaryID = 21
        Case "*": BinaryID = 22
        Case "/": BinaryID = 23
        Case "<": BinaryID = 24
        Case "<=": BinaryID = 25
        Case ">": BinaryID = 26
        Case ">=": BinaryID = 27
        Case "=", "==": BinaryID = 28
        Case "!=": BinaryID = 29
        Case "%": BinaryID = 30
        Case "^": BinaryID = 31
        Case "&": BinaryID = 32
        Case "|": BinaryID = 33
        Case "<<": BinaryID = 34
        Case ">>": BinaryID = 35
        Case Else: Reject "binary operator: " & symbol
    End Select
End Function

Private Sub Expression(ByVal node As ASF_Map)
    If node Is Nothing Then Literal Empty: Exit Sub
    Dim symbol As String, jump As Long, finish As Long, value As Variant, id As Long
    Dim left As ASF_Map, right As ASF_Map, callee As ASF_Map, args As Collection, arg As ASF_Map
    Dim keys As Collection, pair As Collection, spreadCount As Long
    If scoped Then
        Select Case CStr(node.GetValue("type"))
            Case "Literal", "Variable", "Binary", "Unary", "Ternary"
            Case "Call"
                If EmitScopedCall(node) Then Exit Sub
                If node.Exists("callee") Then
                    Set callee = node.GetValue("callee")
                    If callee.GetValue("type") = "Member" Then
                        id = ArrayServiceID(CStr(callee.GetValue("prop")))
                        If id > 0 Then Emit 75, Pool(node), id: Exit Sub
                    End If
                End If
                Emit 67, Pool(node): Exit Sub
            Case Else: Emit 67, Pool(node): Exit Sub
        End Select
    End If
    If optimise Then
        If Fold(node, value) Then Literal value: folded = folded + 1: Exit Sub
    End If
    Select Case CStr(node.GetValue("type"))
        Case "Literal": Literal node.GetValue("value")
        Case "Variable": Emit 2, NameID(CStr(node.GetValue("name")))
        Case "Binary"
            symbol = node.GetValue("op")
            Set left = node.GetValue("left"): Set right = node.GetValue("right")
            Expression left
            If scoped Then Emit 72
            Select Case symbol
                Case "&&", "||"
                    jump = Emit(IIf(symbol = "&&", 8, 9))
                    Expression right: Emit 10
                    finish = Emit(7): a(jump) = used
                    Literal (symbol = "||"): a(finish) = used
                Case "??"
                    jump = Emit(11): Expression right: a(jump) = used
                Case Else
                    Expression right: Emit BinaryID(symbol)
            End Select
        Case "Unary"
            symbol = node.GetValue("op"): Expression node.GetValue("expr")
            Select Case symbol
                Case "!": Emit 12
                Case "-": Emit 13
                Case "typeof": Emit 14
                Case Else: Reject "unary operator: " & symbol
            End Select
        Case "Ternary"
            Expression node.GetValue("cond")
            If scoped Then Emit 72
            jump = Emit(8)
            Expression node.GetValue("trueExpr"): finish = Emit(7)
            a(jump) = used: Expression node.GetValue("falseExpr"): a(finish) = used
        Case "Array"
            Set args = node.GetValue("items")
            Set keys = New Collection
            For Each arg In args
                If arg.GetValue("type") = "SpreadElement" Then
                    spreadCount = spreadCount + 1: keys.Add True
                    Expression arg.GetValue("argument")
                Else
                    keys.Add False
                End If
            Next arg
            For Each arg In args
                If arg.GetValue("type") = "SpreadElement" Then Expression arg.GetValue("argument") Else Expression arg
            Next arg
            If spreadCount > 0 Then Emit 66, Pool(keys), spreadCount Else Emit 50, args.Count
        Case "Object"
            Set args = node.GetValue("items"): Set keys = New Collection
            For Each pair In args
                keys.Add CStr(pair(1))
                Set arg = pair(2)
                If CStr(pair(1)) = "__SPREAD__" Then Expression arg.GetValue("argument") Else Expression arg
            Next pair
            Emit 51, Pool(keys), args.Count
        Case "Template"
            Literal vbNullString
            Set args = node.GetValue("parts")
            For Each arg In args
                If Not arg Is Nothing Then
                    Expression arg: Emit 62, IIf(arg.GetValue("type") = "Literal", 1, 0): Emit 32
                End If
            Next arg
        Case "Index"
            Expression node.GetValue("base"): Expression node.GetValue("index"): Emit 52
        Case "Member"
            Expression node.GetValue("base")
            Emit 53, Pool(CStr(node.GetValue("prop"))), MethodID(CStr(node.GetValue("prop")))
        Case "Call"
            If Not node.Exists("callee") Then Reject "legacy named-call scope"
            Set callee = node.GetValue("callee")
            Set args = node.GetValue("args")
            If CStr(callee.GetValue("type")) = "Member" Then
                symbol = LCase$(CStr(callee.GetValue("prop")))
                ' Only methods without ASF lvalue writeback or callback semantics.
                Select Case symbol
                    Case "length", "trim", "trimstart", "trimend", "touppercase", "tolowercase", "item", "count"
                    Case Else: Reject "method requiring callback/writeback resolution: " & symbol
                End Select
                ' ASF evaluates arguments before the receiver (including side effects).
                For Each arg In args: Expression arg: Next arg
                Expression callee.GetValue("base")
                Emit 54, MethodID(symbol), args.Count
                Exit Sub
            End If
            If CStr(callee.GetValue("type")) <> "Variable" Then Reject "dynamic call/closure"

            symbol = callee.GetValue("name")
            Select Case LCase$(symbol)
                Case "range", "flatten", "clone", "isarray", "isnumeric"
                    For Each arg In args: Expression arg: Next arg
                    If LCase$(symbol) = "clone" Or LCase$(symbol) = "flatten" Then
                        If args.Count > 0 Then Expression args(1)
                    End If
                    Emit 55, BuiltinID(symbol), args.Count
                    Exit Sub
            End Select
            If Not globals.gFuncTable.Exists(symbol) Then Reject "builtin/host call: " & symbol
            id = CLng(globals.gFuncTable.GetValue(symbol))
            On Error Resume Next
            functions.Add id, CStr(id)
            If Err.Number = 0 Then pending.Add id
            Err.Clear: On Error GoTo 0
            Set args = node.GetValue("args")
            For Each arg In args: Expression arg: Next arg
            Emit 5, id, args.Count
        Case Else: Reject "expression: " & CStr(node.GetValue("type"))
    End Select
End Sub

' Folding is deliberately narrow: no locale-sensitive coercion, host calls,
' errors or dead-branch side effects are evaluated during compilation.
Private Function Fold(ByVal node As ASF_Map, ByRef value As Variant) As Boolean
    On Error GoTo NotFolded
    If CStr(node.GetValue("type")) <> "Binary" Then Exit Function
    Dim left As ASF_Map, right As ASF_Map, x As Variant, y As Variant
    Set left = node.GetValue("left"): Set right = node.GetValue("right")
    If left.GetValue("type") <> "Literal" Or right.GetValue("type") <> "Literal" Then Exit Function
    x = left.GetValue("value"): y = right.GetValue("value")
    If VarType(x) <> vbDouble Or VarType(y) <> vbDouble Then Exit Function
    Select Case CStr(node.GetValue("op"))
        Case "+": value = x + y
        Case "-": value = x - y
        Case "*": value = x * y
        Case "<": value = (x < y)
        Case "<=": value = (x <= y)
        Case ">": value = (x > y)
        Case ">=": value = (x >= y)
        Case Else: Exit Function
    End Select
    Fold = True
NotFolded:
    Err.Clear
End Function
Private Sub Condition(ByVal node As ASF_Map, ByVal target As Long)
    If optimise Then
        If node.GetValue("type") = "Binary" Then
            If node.GetValue("op") = "<" Then
                Expression node.GetValue("left"): Expression node.GetValue("right")
                Emit 40, target: fused = fused + 1: Exit Sub
            End If
        End If
    End If
    Expression node: Emit 8, target
End Sub
Private Sub CompileBlock(ByVal block As Collection, ByVal exitLoop As Long, ByVal nextLoop As Long)
    Dim node As ASF_Map
    For Each node In block: Statement node, exitLoop, nextLoop: Next node
End Sub
Private Sub IgnoredStatement(ByVal node As ASF_Map)
    If node Is Nothing Then Exit Sub
    Dim start As Long, i As Long
    If scoped Then
        Emit 68, Pool(node), 1
        Exit Sub
    End If
    If node.GetValue("type") <> "Assign" And node.GetValue("type") <> "ExprStmt" Then Reject "for initializer/step control statement"
    start = used
    Statement node, -1, -1
    For i = start To used - 1: handlers(i) = used: Next i
End Sub
Private Sub Statement(ByVal node As ASF_Map, ByVal exitLoop As Long, ByVal nextLoop As Long)
    Dim kind As String: kind = node.GetValue("type")
    Dim left As ASF_Map, right As ASF_Map, part As ASF_Map
    Dim start As Long, jump As Long, finish As Long, breakPC As Long, continuePC As Long, i As Long
    Dim conds As Collection, blocks As Collection, exits As Collection, item As Variant
    Dim pair As Collection, slot As Long, savedHandler As Long, tryStart As Long
    If scoped Then
        Select Case kind
            Case "Block", "Assign", "ExprStmt", "Return", "If", "While", "Break", "Continue", "TryCatch", "Print"
            Case "For"
                If node.Exists("kind") Then
                    If node.GetValue("kind") <> "classic" Then Emit 68, Pool(node): Exit Sub
                End If
            Case Else: Emit 68, Pool(node): Exit Sub
        End Select
    End If
    Select Case kind
        Case "Block": CompileBlock node.GetValue("body"), -2, -2
        Case "Assign"
            Set left = node.GetValue("left"): Set right = node.GetValue("right")
            If optimise And left.GetValue("type") = "Variable" Then
                If right.GetValue("type") = "Binary" Then
                    If right.GetValue("op") = "+" Then
                        Set part = right.GetValue("left")
                        If part.GetValue("type") = "Variable" Then
                            If part.GetValue("name") = left.GetValue("name") Then
                                Set part = right.GetValue("right")
                                If part.GetValue("type") = "Literal" Then
                                    If VarType(part.GetValue("value")) = vbDouble Then
                                        If part.GetValue("value") = 1 Then
                                            Emit 42, NameID(CStr(left.GetValue("name")))
                                            fused = fused + 1: Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            Expression right
            If left.GetValue("type") = "Variable" Then
                Emit 3, NameID(CStr(left.GetValue("name")))
            ElseIf scoped Then
                Emit 69, Pool(left)
            Else
                CompileWrite left
            End If
        Case "Print"
            Set blocks = node.GetValue("args")
            For Each part In blocks: Expression part: Next part
            Emit 65, blocks.Count
        Case "ExprStmt": Expression node.GetValue("expr"): Emit 4
        Case "Return": Expression node.GetValue("expr"): Emit 6
        Case "Break"
            If exitLoop = -2 Then Exit Sub
            If exitLoop < 0 Then
                If scoped Then Emit 71, 1: Exit Sub
                Reject "break outside loop"
            End If
            Emit 7, exitLoop
        Case "Continue"
            If nextLoop = -2 Then Exit Sub
            If nextLoop < 0 Then
                If scoped Then Emit 71, 2: Exit Sub
                Reject "continue outside loop"
            End If
            Emit 7, nextLoop
        Case "If"
            Set exits = New Collection
            Expression node.GetValue("cond"): jump = Emit(8)
            CompileBlock node.GetValue("then"), exitLoop, nextLoop
            exits.Add Emit(7): a(jump) = used
            Set conds = node.GetValue("elseif_conds"): Set blocks = node.GetValue("elseif_blocks")
            For i = 1 To conds.Count
                Expression conds(i): jump = Emit(8)
                CompileBlock blocks(i), exitLoop, nextLoop
                exits.Add Emit(7): a(jump) = used
            Next i
            If node.GetValue("hasElse") Then CompileBlock node.GetValue("else"), exitLoop, nextLoop
            For Each item In exits: a(CLng(item)) = used: Next item
        Case "ArrayDestructuring"
            scratchCount = scratchCount + 1: slot = NameID("#destruct" & scratchCount)
            Expression node.GetValue("source"): Emit 43, slot
            Set blocks = node.GetValue("targets")
            For i = 1 To blocks.Count
                Set part = blocks(i)
                Emit 2, slot: Emit 60, i - 1: Emit 3, NameID(CStr(part.GetValue("name")))
            Next i
            If node.Exists("restTarget") Then
                Emit 2, slot: Emit 61, blocks.Count: Emit 3, NameID(CStr(node.GetValue("restTarget")))
            End If
        Case "TryCatch"
            ' Reference try/catch consumes break/continue and ignores catch binding.
            savedHandler = currentHandler
            jump = Emit(7): breakPC = Emit(7): a(jump) = used
            currentHandler = breakPC
            CompileBlock node.GetValue("try"), IIf(scoped, -2, -1), IIf(scoped, -2, -1)
            currentHandler = savedHandler
            finish = Emit(7): a(breakPC) = used
            CompileBlock node.GetValue("catch"), IIf(scoped, -2, -1), IIf(scoped, -2, -1)
            a(finish) = used
        Case "Switch"
            scratchCount = scratchCount + 1: slot = NameID("#switch" & scratchCount)
            Expression node.GetValue("expr"): Emit 43, slot
            jump = Emit(7): breakPC = Emit(7): a(jump) = used
            Set blocks = node.GetValue("cases")
            For Each pair In blocks
                Expression pair(1): Emit 2, slot: Emit 41: finish = Emit(8)
                CompileBlock pair(2), breakPC, -1
                Emit 7, breakPC: a(finish) = used
            Next pair
            Set blocks = node.GetValue("default")
            If Not blocks Is Nothing Then CompileBlock blocks, breakPC, -1
            a(breakPC) = used
        Case "While", "For"
            If kind = "For" Then
                If node.Exists("kind") Then
                    If node.GetValue("kind") <> "classic" Then
                        CompileEach node
                        Exit Sub
                    End If
                End If
                IgnoredStatement node.GetValue("init")
            End If
            jump = Emit(7): breakPC = Emit(7): continuePC = Emit(7)
            a(jump) = used: start = used
            Set part = node.GetValue("cond")
            savedHandler = currentHandler
            If scoped And kind = "For" Then
                If savedHandler >= 0 Then currentHandler = -savedHandler - 3 Else currentHandler = -2
            End If
            If Not part Is Nothing Then Condition part, breakPC
            currentHandler = savedHandler
            CompileBlock node.GetValue("body"), breakPC, continuePC
            a(continuePC) = used
            If kind = "For" Then IgnoredStatement node.GetValue("step")
            Emit 7, start: a(breakPC) = used
        Case Else: Reject "statement: " & kind
    End Select
End Sub

Private Sub CompileEach(ByVal node As ASF_Map)
    Dim target As ASF_Map: Set target = node.GetValue("each")
    If target.GetValue("type") <> "Variable" Then Reject "for-in/of non-variable binding"
    scratchCount = scratchCount + 1
    Dim listID As Long, indexID As Long, testPC As Long, exitPC As Long, nextPC As Long, jump As Long
    listID = NameID("#iter" & scratchCount): indexID = NameID("#index" & scratchCount)
    Expression node.GetValue("iter"): Emit 56, IIf(node.GetValue("kind") = "in", 1, 0)
    Emit 43, listID: Literal 0: Emit 43, indexID
    jump = Emit(7): exitPC = Emit(7): nextPC = Emit(7): a(jump) = used
    testPC = used
    Emit 2, indexID: Emit 2, listID: Emit 57: Emit 24: Emit 8, exitPC
    Emit 2, listID: Emit 2, indexID: Emit 58: Emit 3, NameID(CStr(target.GetValue("name")))
    CompileBlock node.GetValue("body"), exitPC, nextPC
    a(nextPC) = used: Emit 2, indexID: Literal 1: Emit 20: Emit 43, indexID
    Emit 7, testPC: a(exitPC) = used
End Sub

Private Sub CompileWrite(ByVal node As ASF_Map)
    Dim chain As New Collection, parts As New Collection, cursor As ASF_Map, part As ASF_Map, i As Long, count As Long
    Set cursor = node
    Do While cursor.GetValue("type") = "Member" Or cursor.GetValue("type") = "Index"
        chain.Add cursor: Set cursor = cursor.GetValue("base")
    Loop
    If cursor.GetValue("type") <> "Variable" Then Reject "assignment through computed receiver"
    Dim name As String: name = cursor.GetValue("name")
    If Left$(name, 1) = "$" Then Reject "assignment into placeholder"
    Emit 2, NameID(name)
    For i = chain.Count To 1 Step -1
        Set part = chain(i)
        If part.GetValue("type") = "Member" Then
            parts.Add Array("Member", CStr(part.GetValue("prop")))
        Else
            parts.Add Array("Index", Empty)
            Expression part.GetValue("index"): count = count + 1
        End If
    Next i
    Emit 64, Pool(parts), count
    Emit 3, NameID(name)
End Sub

Private Function MethodID(ByVal name As String) As Long
    Select Case LCase$(name)
        Case "length": MethodID = 1
        Case "trim": MethodID = 2
        Case "trimstart": MethodID = 3
        Case "trimend": MethodID = 4
        Case "touppercase": MethodID = 5
        Case "tolowercase": MethodID = 6
        Case "item": MethodID = 7
        Case "count": MethodID = 8
    End Select
End Function
Private Function BuiltinID(ByVal name As String) As Long
    Select Case LCase$(name)
        Case "range": BuiltinID = 1
        Case "flatten": BuiltinID = 2
        Case "clone": BuiltinID = 3
        Case "isarray": BuiltinID = 4
        Case "isnumeric": BuiltinID = 5
    End Select
End Function

' Scope-bearing code shares the reference frame objects with escaped closures,
' classes and host evaluators. The program still executes compiled instructions;
' semantic services re-enter this compiler for every child expression/body.
Friend Function BuildNode(ByVal state As ASF_Globals, ByVal node As ASF_Map, ByVal expression As Boolean) As Boolean
    On Error GoTo Unsupported
    scoped = True: optimise = False: Set globals = state: currentHandler = -1
    ReDim op(0 To 127): ReDim a(0 To 127): ReDim b(0 To 127): ReDim handlers(0 To 127)
    ReDim constants(0 To 31): ReDim names(0 To 31)
    ReDim entries(0 To 0): ReDim parameters(0 To 0)
    If expression Then
        ExpressionNode node
        Emit 73
    Else
        Statement node, -1, -1
        Emit 0
    End If
    ReDim Preserve op(0 To used - 1): ReDim Preserve a(0 To used - 1)
    ReDim Preserve b(0 To used - 1): ReDim Preserve handlers(0 To used - 1)
    ReDim Preserve names(0 To IIf(nameUsed = 0, 0, nameUsed - 1))
    BuildNode = True
    Exit Function
Unsupported:
    rejection = Err.Description: Err.Clear
End Function
Private Sub ExpressionNode(ByVal node As ASF_Map)
    Expression node
End Sub

Private Function EmitScopedCall(ByVal node As ASF_Map) As Boolean
    If node.Exists("name") Or Not node.Exists("callee") Then Exit Function
    Dim callee As ASF_Map: Set callee = node.GetValue("callee")
    Dim memberCall As Boolean
    memberCall = (callee.GetValue("type") = "Member")
    If memberCall Then
        Dim receiver As ASF_Map: Set receiver = callee.GetValue("base")
        If receiver.GetValue("type") <> "Variable" Then Exit Function
    ElseIf callee.GetValue("type") <> "Variable" Then
        Exit Function
    End If
    If Not memberCall Then
    Select Case LCase$(CStr(callee.GetValue("name")))
        Case "range", "flatten", "clone", "isarray", "isnumeric", "foreach", "regex", "cwd", "scwd": Exit Function
    End Select
    End If
    Dim args As Collection: Set args = node.GetValue("args")
    Dim arg As ASF_Map
    For Each arg In args
        If Not arg Is Nothing Then
            If arg.GetValue("type") = "SpreadElement" Then Exit Function
        End If
    Next arg
    For Each arg In args: Expression arg: Next arg
    If memberCall Then
        Emit 76, Pool(Array(node, ArrayServiceID(CStr(callee.GetValue("prop"))))), args.Count
    Else
        Emit 74, Pool(node), args.Count
    End If
    EmitScopedCall = True
End Function

Private Function ArrayServiceID(ByVal name As String) As Long
    Select Case LCase$(name)
        Case "map": ArrayServiceID = 1
        Case "filter": ArrayServiceID = 2
        Case "reduce": ArrayServiceID = 3
        Case "slice": ArrayServiceID = 4
        Case "push": ArrayServiceID = 5
        Case "pop": ArrayServiceID = 6
        Case "foreach": ArrayServiceID = 7
        Case "unique": ArrayServiceID = 8
        Case "concat": ArrayServiceID = 9
        Case "shift": ArrayServiceID = 10
        Case "splice": ArrayServiceID = 11
        Case "at": ArrayServiceID = 12
        Case "copywithin": ArrayServiceID = 13
        Case "entries": ArrayServiceID = 14
        Case "every": ArrayServiceID = 15
        Case "from": ArrayServiceID = 16
        Case "of": ArrayServiceID = 17
        Case "reverse": ArrayServiceID = 18
        Case "toreversed": ArrayServiceID = 19
        Case "some": ArrayServiceID = 20
        Case "with": ArrayServiceID = 21
    End Select
End Function
