VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ASF_ScopeVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A trampoline for scope-bearing bytecode. Direct dynamic/named closure calls
' push these return records, never recurse through the VBA interpreter stack.
Friend Sub Execute(ByVal machine As ASF_FastVM, ByVal scope As ASF_ScopeStack, ByVal host As ASF_VM, ByRef result As Variant, ByRef signal As String)
    Dim op() As Long, a() As Long, b() As Long, errors() As Long, names() As String, pool() As Variant
    Dim cells() As ASF_Map, positions() As Long, width As Long, data As Variant
    Dim stack() As Variant, capacity As Long, sp As Long, pc As Long, instruction As Long
    Dim i As Long, id As Long, x As Variant, y As Variant, value As Variant, node As ASF_Map
    Dim savedMachines() As ASF_FastVM, savedScopes() As ASF_ScopeStack
    Dim returnPC() As Long, returnSP() As Long, traces() As String, depth As Long, frames As Long
    Dim child As ASF_FastVM, childScope As ASF_ScopeStack, args As Collection, argBase As Long, trace As String
    Dim errorNumber As Long, errorSource As String, errorDescription As String, entered As Boolean, methodID As Long, descriptor As Variant, target As Long, failedReturn As Boolean
    capacity = 128: frames = 32
    ReDim stack(0 To capacity - 1)
    ReDim savedMachines(0 To frames - 1): ReDim savedScopes(0 To frames - 1)
    ReDim returnPC(0 To frames - 1): ReDim returnSP(0 To frames - 1): ReDim traces(0 To frames - 1)
    On Error GoTo Failed
Activate:
    data = machine.ScopeCode
    op = data(0): a = data(1): b = data(2): pool = data(3): names = data(4): width = data(5): errors = data(6)
    ReDim cells(0 To width - 1): ReDim positions(0 To width - 1)
Dispatch:
    If sp + 2 >= capacity Then capacity = capacity * 2: ReDim Preserve stack(0 To capacity - 1)
    instruction = op(pc)
    Select Case instruction
        Case 0
            value = Empty
            If depth > 0 Then GoTo ReturnFrame
            Exit Sub
        Case 1: CopyValue stack(sp), pool(a(pc)): sp = sp + 1
        Case 2
            id = a(pc)
            If cells(id) Is Nothing Then scope.ResolveCell names(id), cells(id), positions(id), False
            value = Empty
            If Not cells(id) Is Nothing Then CopyValue value, cells(id).CellValue(positions(id))
            If IsEmpty(value) Then CopyValue value, host.FastFunctionReference(names(id), scope)
            CopyValue stack(sp), value: sp = sp + 1
        Case 3
            id = a(pc): sp = sp - 1
            If Left$(names(id), 1) <> "$" Then
                If cells(id) Is Nothing Then scope.ResolveCell names(id), cells(id), positions(id), True
                cells(id).CellStore positions(id), stack(sp)
            End If
            stack(sp) = Empty
        Case 4: sp = sp - 1: stack(sp) = Empty
        Case 6
            sp = sp - 1: CopyValue value, stack(sp): stack(sp) = Empty
            scope.SetLocalValue "__return", value
            If depth > 0 Then GoTo ReturnFrame
            CopyValue result, value: signal = "RETURN": Exit Sub
        Case 7: pc = a(pc): GoTo Dispatch
        Case 8, 9
            sp = sp - 1: value = host.ThunkTruthy(stack(sp)): stack(sp) = Empty
            If CBool(value) = (instruction = 9) Then pc = a(pc): GoTo Dispatch
        Case 10: stack(sp - 1) = host.ThunkTruthy(stack(sp - 1))
        Case 11
            If Not IsEmpty(stack(sp - 1)) And Not IsNull(stack(sp - 1)) Then pc = a(pc): GoTo Dispatch
            sp = sp - 1: stack(sp) = Empty
        Case 12: stack(sp - 1) = Not host.ThunkTruthy(stack(sp - 1))
        Case 13: stack(sp - 1) = -CDbl(stack(sp - 1))
        Case 14: value = host.FastType(stack(sp - 1)): CopyValue stack(sp - 1), value
        Case 20 To 35
            sp = sp - 1: CopyValue x, stack(sp - 1): CopyValue y, stack(sp)
            CopyValue value, host.ThunkBinary(instruction - 20, x, y)
            stack(sp) = Empty: CopyValue stack(sp - 1), value
        Case 65
            argBase = sp - a(pc): Set args = New Collection
            For i = argBase To sp - 1: args.Add stack(i): stack(i) = Empty: Next i
            host.FastPrint args: sp = argBase: Set args = Nothing
        Case 75
            Set node = pool(a(pc)): CopyValue value, host.FastScopeCall(node, scope, b(pc))
            CopyValue stack(sp), value: sp = sp + 1
            For i = 0 To width - 1: Set cells(i) = Nothing: Next i
        Case 67
            Set node = pool(a(pc)): CopyValue value, host.FastScopeExpression(node, scope)
            CopyValue stack(sp), value: sp = sp + 1
            For i = 0 To width - 1: Set cells(i) = Nothing: Next i
        Case 68
            Set node = pool(a(pc)): signal = host.FastScopeStatement(node, scope)
            For i = 0 To width - 1: Set cells(i) = Nothing: Next i
            If b(pc) = 1 Then signal = vbNullString
            If signal = "RETURN" And depth > 0 Then CopyValue value, scope.GetValue("__return"): GoTo ReturnFrame
            If signal = "ERR" And errors(pc) >= 0 Then
                target = errors(pc): GoTo CatchError
            End If
            If signal = "ERR" And depth > 0 Then
                host.FastCallError traces(depth), "Error during closure execution"
                value = Empty: failedReturn = True: GoTo ReturnFrame
            End If
            If Len(signal) > 0 Then Exit Sub
        Case 69
            sp = sp - 1: Set node = pool(a(pc)): host.FastScopeAssign node, stack(sp), scope
            stack(sp) = Empty
            For i = 0 To width - 1: Set cells(i) = Nothing: Next i
        Case 71: signal = IIf(a(pc) = 1, "BREAK", "CONTINUE"): Exit Sub
        Case 72: value = stack(sp - 1): CopyValue stack(sp - 1), value
        Case 73: sp = sp - 1: CopyValue result, stack(sp): Exit Sub
        Case 74, 76
            argBase = sp - b(pc): Set args = New Collection
            For i = argBase To sp - 1: args.Add stack(i): stack(i) = Empty: Next i
            sp = argBase
            If instruction = 76 Then
                descriptor = pool(a(pc)): Set node = descriptor(0): methodID = descriptor(1)
            Else
                Set node = pool(a(pc))
            End If
            Set child = Nothing: Set childScope = Nothing: value = Empty
            If instruction = 76 Then
                entered = host.FastEnterMember(node, args, scope, child, childScope, trace, value, methodID)
            Else
                entered = host.FastEnterCall(node, args, scope, child, childScope, trace, value)
            End If
            If entered Then
                depth = depth + 1
                If depth >= frames Then
                    frames = frames * 2
                    ReDim Preserve savedMachines(0 To frames - 1): ReDim Preserve savedScopes(0 To frames - 1)
                    ReDim Preserve returnPC(0 To frames - 1): ReDim Preserve returnSP(0 To frames - 1): ReDim Preserve traces(0 To frames - 1)
                End If
                Set savedMachines(depth) = machine: Set savedScopes(depth) = scope
                returnPC(depth) = pc + 1: returnSP(depth) = argBase: traces(depth) = trace
                Set machine = child: Set scope = childScope: pc = 0: Set args = Nothing
                GoTo Activate
            End If
            CopyValue stack(sp), value: sp = sp + 1: Set args = Nothing
            For i = 0 To width - 1: Set cells(i) = Nothing: Next i
        Case Else: Err.Raise 5, "ASF.ScopeVM", "Invalid scoped opcode " & instruction
    End Select
    value = Empty: x = Empty: y = Empty
    Set node = Nothing
    pc = pc + 1: GoTo Dispatch
ReturnFrame:
    For i = returnSP(depth) To sp - 1: stack(i) = Empty: Next i
    x = Empty: y = Empty: Set node = Nothing: Set args = Nothing
    Erase cells: Erase positions
    host.FastLeaveCall scope, traces(depth), value, failedReturn
    failedReturn = False
    Set machine = savedMachines(depth): Set savedMachines(depth) = Nothing
    Set scope = savedScopes(depth): Set savedScopes(depth) = Nothing
    pc = returnPC(depth): sp = returnSP(depth): depth = depth - 1
    CopyValue stack(sp), value: sp = sp + 1: signal = vbNullString
    GoTo Activate
Failed:
    errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
    Err.Clear: Resume Recover
Recover:
    target = errors(pc)
    If target <= -3 Then target = -target - 3
    If target >= 0 Then
        If errors(pc) >= 0 Then host.FastStatementError errorDescription
        GoTo CatchError
    End If
    If depth > 0 Then
        If errors(pc) <> -2 Then host.FastStatementError errorDescription
        host.FastCallError traces(depth), "Error during closure execution"
        value = Empty: failedReturn = True: GoTo ReturnFrame
    End If
    If errors(pc) = -2 Then signal = "ERR": Exit Sub
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorDescription
    Exit Sub
CatchError:
    argBase = returnSP(depth)
    For i = argBase To sp - 1: stack(i) = Empty: Next i
    sp = argBase: pc = target: signal = vbNullString
    value = Empty: x = Empty: y = Empty: Set node = Nothing: Set args = Nothing
    For i = 0 To width - 1: Set cells(i) = Nothing: Next i
    GoTo Dispatch
End Sub
Private Sub CopyValue(ByRef target As Variant, ByRef source As Variant)
    If IsObject(source) Then Set target = source Else target = source
End Sub
