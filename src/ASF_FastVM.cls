VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ASF_FastVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Immutable instruction arrays; all stacks, bindings and return frames are local
' to Execute. A function call changes pc, not the VBA call stack.
Private op() As Long, a() As Long, b() As Long, handlers() As Long
Private constants() As Variant, names() As String, entries() As Long, parameters() As Variant
Private width As Long, supported As Boolean, rejection As String
Private foldCount As Long, fusionCount As Long
Private positions() As Long, writable() As Boolean, argumentID() As Long
Private optimised As Boolean

Friend Sub Prepare(ByVal globals As ASF_Globals, ByVal index As Long, ByVal optimize As Boolean)
    Dim compiler As New ASF_Bytecode, data As Variant
    optimised = optimize
    supported = compiler.Build(globals, index, optimize)
    rejection = compiler.Reason
    If Not supported Then Exit Sub
    data = compiler.Export
    op = data(0): a = data(1): b = data(2): handlers = data(3)
    constants = data(4): names = data(5): entries = data(6): parameters = data(7)
    width = data(8): foldCount = data(9): fusionCount = data(10)
    If width = 0 Then width = 1
    ReDim positions(0 To UBound(op))
    ReDim writable(0 To width - 1): ReDim argumentID(0 To width - 1)
    Dim i As Long
    For i = 0 To UBound(op): positions(i) = -1: Next i
    For i = 0 To UBound(names)
        writable(i) = (Left$(names(i), 1) <> "$")
        If Not writable(i) Then
            argumentID(i) = Val(Mid$(names(i), 2))
            If names(i) <> "$" & CStr(argumentID(i)) Then argumentID(i) = 0
        End If
    Next i
End Sub
Friend Property Get Eligible() As Boolean
    Eligible = supported
End Property
Friend Property Get Reason() As String
    Reason = rejection
End Property
Friend Property Get FoldedConstants() As Long
    FoldedConstants = foldCount
End Property
Friend Property Get FusedInstructions() As Long
    FusedInstructions = fusionCount
End Property

Friend Function Execute(ByRef result As Variant, ByVal arguments As Variant, ByVal host As ASF_VM) As Boolean
    Dim stack() As Variant, values() As Variant, binding() As Long, shadows() As Long
    Dim returnPC() As Long, returnSP() As Long, owned() As Boolean
    Dim frames As Long, stackCapacity As Long, depth As Long, base As Long
    Dim pc As Long, sp As Long, id As Long, target As Long, i As Long, argc As Long, argBase As Long
    Dim paramIDs As Variant, value As Variant, x As Variant, y As Variant, instruction As Long
    Dim items As Collection, map As ASF_Map, other As ASF_Map, arr As Variant, key As Variant, j As Long, extra As Long
    frames = 32: stackCapacity = 128
    ReDim stack(0 To stackCapacity - 1): ReDim values(0 To frames * width - 1)
    ReDim binding(0 To width - 1): ReDim shadows(0 To frames * width - 1)
    ReDim owned(0 To frames * width - 1)
    ReDim returnPC(0 To frames - 1): ReDim returnSP(0 To frames - 1)
    For i = 0 To width - 1: binding(i) = -1: Next i
    For i = 0 To UBound(names)
        If names(i) = "__option_base" Then
            binding(i) = i: values(i) = 1
        ElseIf argumentID(i) > 0 Then
            id = argumentID(i)
            If id > 0 And IsArray(arguments) Then
                If id <= UBound(arguments) - LBound(arguments) + 1 Then
                    binding(i) = i: CopyValue values(i), arguments(LBound(arguments) + id - 1)
                End If
            End If
        End If
    Next i
    On Error GoTo RuntimeFailure
Dispatch:
    If sp + 2 >= stackCapacity Then
        stackCapacity = stackCapacity * 2
        ReDim Preserve stack(0 To stackCapacity - 1)
    End If
    instruction = op(pc)
    Select Case instruction
        Case 0: Exit Function
        Case 1
            CopyValue stack(sp), constants(a(pc)): sp = sp + 1
        Case 2
            target = binding(a(pc))
            If target < 0 Then stack(sp) = Empty Else CopyValue stack(sp), values(target)
            sp = sp + 1
        Case 3
            id = a(pc): sp = sp - 1
            If writable(id) Then
                target = binding(id)
                If target < 0 Then
                    target = base + id: binding(id) = target
                    shadows(target) = -1: owned(target) = True
                End If
                CopyValue values(target), stack(sp)
            End If
            stack(sp) = Empty
        Case 4
            sp = sp - 1: stack(sp) = Empty
        Case 5 ' explicit CALL frame
            argc = b(pc): argBase = sp - argc: depth = depth + 1
            If depth >= frames Then
                frames = frames * 2
                ReDim Preserve values(0 To frames * width - 1)
                ReDim Preserve shadows(0 To frames * width - 1)
                ReDim Preserve owned(0 To frames * width - 1)
                ReDim Preserve returnPC(0 To frames - 1): ReDim Preserve returnSP(0 To frames - 1)
            End If
            base = depth * width: returnPC(depth) = pc + 1: returnSP(depth) = argBase
            paramIDs = parameters(a(pc))
            For i = 1 To paramIDs(0)
                id = paramIDs(i): target = base + id
                If Not owned(target) Then shadows(target) = binding(id): owned(target) = True
                binding(id) = target
                If i <= argc Then CopyValue values(target), stack(argBase + i - 1) Else values(target) = Empty
            Next i
            id = paramIDs(paramIDs(0) + 1)
            If id >= 0 Then
                target = base + id
                If Not owned(target) Then shadows(target) = binding(id): owned(target) = True
                binding(id) = target
                j = paramIDs(paramIDs(0) + 2): extra = argc - j
                If extra = 1 Then
                    If IsEmpty(stack(argBase + j)) Then extra = 0
                End If
                If extra <= 0 Then
                    ReDim arr(-1 To -1)
                Else
                    ReDim arr(1 To extra)
                    For i = 1 To extra: CopyValue arr(i), stack(argBase + j + i - 1): Next i
                End If
                values(target) = arr
            End If
            For i = argBase To sp - 1: stack(i) = Empty: Next i
            sp = argBase: pc = entries(a(pc)): GoTo Dispatch
        Case 6 ' RET
            sp = sp - 1: CopyValue value, stack(sp): stack(sp) = Empty
            If depth = 0 Then
                CopyValue result, value: Execute = True: Exit Function
            End If
ReturnToCaller:
            For i = 0 To width - 1
                target = base + i
                If owned(target) Then
                    binding(i) = shadows(target): owned(target) = False
                End If
                values(target) = Empty
            Next i
            pc = returnPC(depth): sp = returnSP(depth)
            depth = depth - 1: base = depth * width
            CopyValue stack(sp), value: value = Empty: sp = sp + 1
            GoTo Dispatch
        Case 7: pc = a(pc): GoTo Dispatch
        Case 8, 9
            sp = sp - 1: value = Truth(stack(sp)): stack(sp) = Empty
            If CBool(value) = (instruction = 9) Then pc = a(pc): GoTo Dispatch
        Case 10: stack(sp - 1) = Truth(stack(sp - 1))
        Case 11
            If Not IsEmpty(stack(sp - 1)) And Not IsNull(stack(sp - 1)) Then pc = a(pc): GoTo Dispatch
            sp = sp - 1: stack(sp) = Empty
        Case 12: stack(sp - 1) = Not Truth(stack(sp - 1))
        Case 13: stack(sp - 1) = -CDbl(stack(sp - 1))
        Case 14: CopyValue value, host.FastType(stack(sp - 1)): CopyValue stack(sp - 1), value
        Case 20 To 35, 40, 41
            sp = sp - 1: CopyValue x, stack(sp - 1): CopyValue y, stack(sp)
            Select Case instruction
                Case 20
                    If VarType(x) = vbString Or VarType(y) = vbString Then
                        value = IIf(IsNull(x), x, CStr(x)) & IIf(IsNull(y), y, CStr(y))
                    Else
                        value = x + y
                    End If
                Case 21: value = x - y
                Case 22: value = x * y
                Case 23
                    If y = 0 Then Err.Raise vbObjectError + 2001, "VM.EvalExprNode", "Division by zero"
                    value = x / y
                Case 41: value = (x = y)
                Case 24, 40: value = (x < y)
                Case 25: value = (x <= y)
                Case 26: value = (x > y)
                Case 27: value = (x >= y)
                Case 28: value = (IIf(IsNull(x), Empty, x) = IIf(IsNull(y), Empty, y))
                Case 29: value = (x <> y)
                Case 30: value = x Mod y
                Case 31: value = x ^ y
                Case 32: value = CStr(x) & CStr(y)
                Case 33: value = x Or y
                Case 34, 35
                    value = Empty
                    If IsNumeric(x) And IsNumeric(y) Then
                        If instruction = 34 Then value = Floor(Floor(x) * (2 ^ Abs(Floor(y)))) Else value = Floor(Floor(x) / (2 ^ Abs(Floor(y))))
                    End If
            End Select
            stack(sp) = Empty
            If instruction = 40 Then
                sp = sp - 1: stack(sp) = Empty
                If Not Truth(value) Then pc = a(pc): GoTo Dispatch
            Else
                stack(sp - 1) = value
            End If
        Case 43 ' compiler temporaries are frame-local, never caller bindings
            id = a(pc): target = base + id: sp = sp - 1
            If Not owned(target) Then shadows(target) = binding(id): owned(target) = True
            binding(id) = target: CopyValue values(target), stack(sp): stack(sp) = Empty
        Case 42 ' fused variable + Double(1), preserving SafeAdd coercion
            id = a(pc): target = binding(id): x = Empty
            If target >= 0 Then CopyValue x, values(target)
            If VarType(x) = vbString Then value = CStr(x) & "1" Else value = x + 1#
            If writable(id) Then
                If target < 0 Then
                    target = base + id: binding(id) = target: shadows(target) = -1: owned(target) = True
                End If
                CopyValue values(target), value
            End If
        Case 50 ' array literal
            argc = a(pc): argBase = sp - argc
            If argc = 0 Then ReDim arr(-1 To -1) Else ReDim arr(1 To argc)
            For i = 1 To argc: CopyValue arr(i), stack(argBase + i - 1): Next i
            For i = argBase To sp - 1: stack(i) = Empty: Next i
            sp = argBase: stack(sp) = arr: sp = sp + 1
        Case 51 ' object literal
            Set items = constants(a(pc)): Set map = New ASF_Map
            argBase = sp - b(pc)
            For i = 1 To items.Count
                CopyValue x, stack(argBase + i - 1)
                If CStr(items(i)) <> "__SPREAD__" Then
                    map.SetValue CStr(items(i)), x
                ElseIf TypeName(x) = "ASF_Map" Then
                    Set other = x
                    For Each key In other.Keys: map.SetValue CStr(key), other.GetValue(CStr(key)): Next key
                    Set other = Nothing
                ElseIf IsArray(x) Then
                    For j = LBound(x) To UBound(x): map.SetValue CStr(j - LBound(x) + 1), x(j): Next j
                ElseIf VarType(x) = vbString Then
                    For j = 1 To Len(x): map.SetValue CStr(j), Mid$(x, j, 1): Next j
                End If
            Next i
            For i = argBase To sp - 1: stack(i) = Empty: Next i
            sp = argBase: Set stack(sp) = map: sp = sp + 1: Set map = Nothing: Set items = Nothing
        Case 52
            sp = sp - 1: CopyValue x, stack(sp - 1): id = CLng(stack(sp))
            If id < 1 Then Err.Raise 5
            If IsArray(x) Then
                CopyValue value, x(LBound(x) + id - 1)
            ElseIf TypeName(x) = "ASF_Map" Then
                Set map = x: CopyValue value, map.GetValue(CStr(id)): Set map = Nothing
            Else
                Err.Raise 5
            End If
            stack(sp) = Empty: CopyValue stack(sp - 1), value
        Case 53
            CopyValue x, stack(sp - 1)
            If TypeName(x) = "ASF_Map" Then
                Set map = x
                If map.Exists("classMethods") Then
                    Set items = New Collection
                    CopyValue value, host.FastMember(x, CStr(constants(a(pc))), items, False)
                    Set items = Nothing
                ElseIf optimised Then
                    CopyValue value, map.FastGetValue(CStr(constants(a(pc))), positions(pc))
                Else
                    CopyValue value, map.GetValue(CStr(constants(a(pc))))
                End If
                Set map = Nothing
            ElseIf VarType(x) = vbString And b(pc) >= 1 And b(pc) <= 6 Then
                value = StringPrimitive(x, b(pc))
            ElseIf IsArray(x) And b(pc) = 1 Then
                value = 0
                If UBound(x) <> -1 Then value = UBound(x) - LBound(x) + 1
            Else
                Set items = New Collection
                CopyValue value, host.FastMember(x, CStr(constants(a(pc))), items, False)
                Set items = Nothing
            End If
            CopyValue stack(sp - 1), value
        Case 54
            argc = b(pc): sp = sp - 1: CopyValue x, stack(sp): stack(sp) = Empty
            argBase = sp - argc: Set items = New Collection
            For i = argBase To sp - 1: items.Add stack(i): stack(i) = Empty: Next i
            If VarType(x) = vbString And a(pc) >= 1 And a(pc) <= 6 Then
                value = StringPrimitive(x, a(pc))
            ElseIf IsArray(x) And a(pc) = 1 Then
                value = 0
                If UBound(x) <> -1 Then value = UBound(x) - LBound(x) + 1
            Else
                CopyValue value, host.FastMember(x, MethodName(a(pc)), items, True)
            End If
            sp = argBase: CopyValue stack(sp), value: sp = sp + 1: Set items = Nothing
        Case 55
            argc = b(pc): extra = 0
            If argc > 0 And (a(pc) = 2 Or a(pc) = 3) Then extra = 1
            argBase = sp - argc - extra: Set items = New Collection
            For i = argBase To sp - 1: items.Add stack(i): stack(i) = Empty: Next i
            CopyValue value, host.FastBuiltin(a(pc), items)
            sp = argBase: CopyValue stack(sp), value: sp = sp + 1: Set items = Nothing
        Case 56 ' materialize the reference's iteration snapshot once
            CopyValue x, stack(sp - 1): Set items = New Collection
            If IsArray(x) Then
                For i = LBound(x) To UBound(x)
                    If a(pc) = 1 Then items.Add 1 + i - LBound(x) Else items.Add x(i)
                Next i
            ElseIf VarType(x) = vbString Then
                For i = 1 To Len(x)
                    If a(pc) = 1 Then items.Add i Else items.Add Mid$(x, i, 1)
                Next i
            ElseIf TypeName(x) = "ASF_Map" Then
                Set map = x
                For Each key In map.Keys
                    If a(pc) = 1 Then items.Add key Else items.Add map.GetValue(CStr(key))
                Next key
                Set map = Nothing
            End If
            Set stack(sp - 1) = items: Set items = Nothing
        Case 57
            Set items = stack(sp - 1): stack(sp - 1) = items.Count: Set items = Nothing
        Case 58
            sp = sp - 1: Set items = stack(sp - 1): id = CLng(stack(sp)) + 1
            CopyValue value, items(id): stack(sp) = Empty: CopyValue stack(sp - 1), value: Set items = Nothing
        Case 60
            CopyValue x, stack(sp - 1)
            If Not IsArray(x) Then Err.Raise 5
            id = LBound(x) + a(pc): value = Empty
            If id <= UBound(x) Then CopyValue value, x(id)
            CopyValue stack(sp - 1), value
        Case 61
            CopyValue x, stack(sp - 1)
            If Not IsArray(x) Then Err.Raise 5
            id = UBound(x) - LBound(x) + 1 - a(pc)
            If id <= 0 Then
                ReDim arr(-1 To -1)
            Else
                ReDim arr(1 To id)
                For i = 1 To id: CopyValue arr(i), x(LBound(x) + a(pc) + i - 1): Next i
            End If
            stack(sp - 1) = arr
        Case 62
            value = host.FastString(stack(sp - 1), a(pc) = 1): stack(sp - 1) = value
        Case 64
            argBase = sp - b(pc) - 2: Set items = New Collection
            For i = argBase + 2 To sp - 1: items.Add stack(i): Next i
            CopyValue x, stack(argBase + 1): CopyValue y, stack(argBase)
            host.FastWrite x, y, constants(a(pc)), items
            For i = argBase To sp - 1: stack(i) = Empty: Next i
            sp = argBase: CopyValue stack(sp), x: sp = sp + 1: Set items = Nothing
        Case 65
            argBase = sp - a(pc): Set items = New Collection
            For i = argBase To sp - 1: items.Add stack(i): stack(i) = Empty: Next i
            host.FastPrint items: sp = argBase: Set items = Nothing
        Case 66
            Set items = constants(a(pc)): argBase = sp - items.Count - b(pc)
            extra = items.Count - b(pc)
            For i = argBase To argBase + b(pc) - 1
                CopyValue x, stack(i)
                If VarType(x) = vbString Then
                    extra = extra + Len(x)
                ElseIf IsArray(x) Then
                    If UBound(x) <> -1 Then extra = extra + UBound(x) - LBound(x) + 1
                ElseIf Not IsEmpty(x) Then
                    extra = extra + 1
                End If
            Next i
            If extra = 0 Then ReDim arr(1 To 1) Else ReDim arr(1 To extra)
            id = 1
            For i = 1 To items.Count
                CopyValue x, stack(argBase + b(pc) + i - 1)
                If Not CBool(items(i)) Then
                    CopyValue arr(id), x: id = id + 1
                ElseIf VarType(x) = vbString Then
                    For j = 1 To Len(x): arr(id) = Mid$(x, j, 1): id = id + 1: Next j
                ElseIf IsArray(x) Then
                    If UBound(x) <> -1 Then
                        For j = LBound(x) To UBound(x): CopyValue arr(id), x(j): id = id + 1: Next j
                    End If
                Else
                    CopyValue arr(id), x: id = id + 1
                End If
            Next i
            For i = argBase To sp - 1: stack(i) = Empty: Next i
            sp = argBase: stack(sp) = arr: sp = sp + 1: Set items = Nothing
        Case Else: Err.Raise 5, "ASF.FastVM", "Unknown opcode"
    End Select
    pc = pc + 1: GoTo Dispatch
RuntimeFailure:
    Err.Clear
    Resume Recover
Recover:
    ' No retry: an error either jumps to its precompiled handler or unwinds this
    ' invocation like InvokeClosure (Empty). Top-level errors leave OUTPUT_ alone.
    target = handlers(pc)
    If target >= 0 Then
        For i = returnSP(depth) To sp - 1: stack(i) = Empty: Next i
        sp = returnSP(depth): pc = target: GoTo Dispatch
    End If
    If depth = 0 Then Exit Function
    For i = returnSP(depth) To sp - 1: stack(i) = Empty: Next i
    value = Empty: GoTo ReturnToCaller
End Function

Private Function Truth(ByRef value As Variant) As Boolean
    If IsObject(value) Then Truth = Not value Is Nothing: Exit Function
    If IsArray(value) Then Truth = True: Exit Function
    If IsNull(value) Or IsEmpty(value) Then Exit Function
    If VarType(value) = vbBoolean Then Truth = CBool(value): Exit Function
    If IsNumeric(value) Then Truth = (CDbl(value) <> 0): Exit Function
    If VarType(value) = vbString Then
        If LCase$(value) = "false" Then Exit Function
        Truth = (value <> vbNullString): Exit Function
    End If
    Truth = True
End Function
Private Function Floor(ByVal value As Variant) As Long
    Dim whole As Double: whole = CDbl(Fix(value))
    Floor = CLng(whole + ((CDbl(value) <> whole) And (CDbl(value) < 0)))
End Function

Private Sub CopyValue(ByRef target As Variant, ByRef source As Variant)
    If IsObject(source) Then Set target = source Else target = source
End Sub

Private Function StringPrimitive(ByRef value As Variant, ByVal id As Long) As Variant
    Select Case id
        Case 1: StringPrimitive = Len(value)
        Case 2: StringPrimitive = Trim$(value)
        Case 3: StringPrimitive = LTrim$(value)
        Case 4: StringPrimitive = RTrim$(value)
        Case 5: StringPrimitive = UCase$(value)
        Case 6: StringPrimitive = LCase$(value)
    End Select
End Function
Private Function MethodName(ByVal id As Long) As String
    Select Case id
        Case 1: MethodName = "length"
        Case 2: MethodName = "trim"
        Case 3: MethodName = "trimstart"
        Case 4: MethodName = "trimend"
        Case 5: MethodName = "touppercase"
        Case 6: MethodName = "tolowercase"
        Case 7: MethodName = "item"
        Case 8: MethodName = "count"
    End Select
End Function
