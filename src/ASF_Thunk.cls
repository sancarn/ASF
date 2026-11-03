VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASF_Thunk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Internal native backend. The only cached state is immutable code and AST
' lowering. Every call (including reentry) owns its frame, stack and temporaries.
Private programs As Collection
Private entries As Double
#If VBA7 And Win64 And Not Mac Then
Private Type NativeFrame
    code As LongPtr
    constants As LongPtr
    stack As LongPtr
    copy As LongPtr
    clear As LongPtr
    count As Long
    capacity As Long
    pc As Long
    sp As Long
    request As Long
    argument As Long
    signal As Long
    error As Long
    constantCount As Long
    reserved As Long
End Type
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" (ByVal address As LongPtr, ByVal size As LongPtr, ByVal allocation As Long, ByVal protection As Long) As LongPtr
Private Declare PtrSafe Function VirtualFree Lib "kernel32" (ByVal address As LongPtr, ByVal size As LongPtr, ByVal freeType As Long) As Long
Private Declare PtrSafe Function VirtualProtect Lib "kernel32" (ByVal address As LongPtr, ByVal size As LongPtr, ByVal protection As Long, ByRef previous As Long) As Long
Private Declare PtrSafe Function FlushInstructionCache Lib "kernel32" (ByVal process As LongPtr, ByVal address As LongPtr, ByVal size As LongPtr) As Long
Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal name As LongPtr) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal module As LongPtr, ByVal name As String) As LongPtr
Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal count As LongPtr)
Private Declare PtrSafe Function NativeInvoke Lib "user32" Alias "CallWindowProcW" (ByVal address As LongPtr, ByRef frame As NativeFrame, ByVal a As LongPtr, ByVal b As LongPtr, ByVal c As LongPtr) As LongPtr
Private entry As LongPtr, copyAddress As LongPtr, clearAddress As LongPtr
#End If

Friend Property Get EntryCount() As Double
    EntryCount = entries
End Property
Friend Sub EnsureAvailable()
#If VBA7 And Win64 And Not Mac Then
    If entry <> 0 Then Exit Sub
    Dim dll As String: dll = "oleaut32.dll"
    Dim module As LongPtr: module = GetModuleHandle(StrPtr(dll))
    copyAddress = GetProcAddress(module, "VariantCopy")
    clearAddress = GetProcAddress(module, "VariantClear")
    If copyAddress = 0 Or clearAddress = 0 Then Err.Raise 5, "ASF.Thunk", "Automation helpers unavailable"
    Dim hexCode As String
' BEGIN GENERATED THUNK
    hexCode = hexCode & "4885c90f84c905000056534889cb4883ec38488b094885c90f847401000048837b10000f846901000048837b08000f845e01000048837b18000f845301000048837b20000f84480100008b5340486343"
    hexCode = hexCode & "3085d20f85d7000000c743380000000085c00f882a01000048be000000000000f07f3943280f8e170100004c6353344585d20f880a0100008b532c4139d20f8ffe000000489848c1e0044803034989c1"
    hexCode = hexCode & "4585d20f84b70000004963ca8b004c8d0449488b4b104a8d4cc1e883f80a0f84bc0200000f8f0e01000083f8060f841d0400000f8ec700000083f8080f842301000083f8090f85930000004885c90f84"
    hexCode = hexCode & "9e0000000fb7116683fa030f842f0400000f872903000031c06683fa02750c31c066837908000f95c0f7d841b90b0000006689410866448909e962010000669085c07e5e3943287c5948c1e004488d4c"
    hexCode = hexCode & "01f083fa010f849501000031c083fa020f84ba0100004883c4385b5ec30f1f008b0083f80a742b0f8f5301000083f80674207e2c83f808741983f809741483f807750f418b410489433085c00f89e0fe"
    hexCode = hexCode & "ffff4883c43848c7c0ffffffff5b5ec383f8010f84870100007e1b83f80674e2418b5104894338b80100000089533c4883c4385b5ec3f7d84819c04883c4385b5ec3660f1f44000083f80e0f84370100"
    hexCode = hexCode & "007e5d83f80f75184885c974a50fb71183e2fb6683fa0974b7e992000000669083f810758d4885c974886683390b758266837908000f95c24531c04183ea0183f8100f94c0664489014489533438c275"
    hexCode = hexCode & "5fe94dffffff662e0f1f84000000000083f80c0f848702000083f80d75624139d20f843bffffff418d4201488b53108943344b8d0452488d04c2668338000f851effffff418b510441bb0b0000006644"
    hexCode = hexCode & "8918f7da6619d2668950080f1f4400008b433083c00189433085c00f89d1fdffffe9ecfeffff662e0f1f84000000000083f80b0f84f7feffffe9d4feffff669083f80e74530f8fc7feffff83f80c0f85"
    hexCode = hexCode & "75ffffffe9b9feffff0f1f80000000008b4108833903752085c0781c894330c7434000000000c7433800000000e966fdffff660f1f44000031c0e947feffff908b410cebce0f1f00418b410489434031"
    hexCode = hexCode & "c0e930feffff662e0f1f8400000000004139d20f8459feffff4963410485c00f884dfeffff3b43480f8d44feffff488d1440488b43084b8d0c52488d14d0488b4310488d0cc8ff531889434485c00f88"
    hexCode = hexCode & "5202000083433401e913ffffff0f1f004183fa010f8e08feffff668379e805418b51040f851bfeffff668339050f8511fefffff20f1041f0f20f1049084989f1660fd64424284c234c242866490f7ec8"
    hexCode = hexCode & "4939f10f84ebfdffff4921f04939f00f84dffdffff83fa050f84b00100000f8f9c00000083fa020f84930100007e6183fa030f845001000031c0660f2fc80f97c00f1f8000000000ba0b000000f7d866"
    hexCode = hexCode & "8951e8668941f04531c04183ea016644890144895334e965feffff0f1f4400006683fa050f84d60000006683fa0b0f84cbfcffffe967fdffff0f1f800000000085d20f841801000083fa010f8553fdff"
    hexCode = hexCode & "fff20f5cc166490f7ec04d89c14921f14939f10f843bfdffff4c8941f0eb989083fa080f84270100007e2583fa090f8520fdffff31c0660f2ec1ba010000000f9ac00f45c2e95effffff660f1f440000"
    hexCode = hexCode & "83fa060f84e700000083fa070f851501000031c0660f2fc10f93c0e938ffffff4885c90f84b9fcffff668339010f87a0fcffffe92fffffff0f1f8400000000004885c90f8499fcffffff532089434485"
    hexCode = hexCode & "c00f88bf000000836b3401e980fdffff31c0660fefd2ba01000000660f2e51080f9ac00f45c2e9eefbffff0f1f440000448b510831c04585d20f95c0e9d8fbffff0f1f8000000000660fefdb660f2ecb"
    hexCode = hexCode & "7a060f845cfcfffff20f5ec166490f7ec0e904ffffff662e0f1f840000000000f20f58c166490f7ec0e9ecfeffff6690f20f59c166490f7ec0e9dcfeffff31c0660f2fc80f93c0e96cfeffff0f1f4000"
    hexCode = hexCode & "31c0660f2fc10f97c0e95afeffff669031c0660f2ec1ba000000000f9bc00f45c2e942feffff48c7c0feffffffe984fbffff4883c8ffc3ba08000000e9d3fbffff909090909090909090909090909090"
' END GENERATED THUNK
    Dim bytes() As Byte, i As Long, size As Long, previous As Long
    size = Len(hexCode) \ 2
    ReDim bytes(0 To size - 1)
    For i = 0 To size - 1
        bytes(i) = CByte("&H" & Mid$(hexCode, i * 2 + 1, 2))
    Next i
    entry = VirtualAlloc(0, size, &H3000, 4) ' read/write, never RWX
    If entry = 0 Then Err.Raise 7, "ASF.Thunk", "Cannot allocate native code"
    MoveMemory entry, VarPtr(bytes(0)), size
    If VirtualProtect(entry, size, &H20, previous) = 0 Then
        VirtualFree entry, 0, &H8000: entry = 0
        Err.Raise 5, "ASF.Thunk", "Cannot protect native code"
    End If
    If FlushInstructionCache(-1, entry, size) = 0 Then
        VirtualFree entry, 0, &H8000: entry = 0
        Err.Raise 5, "ASF.Thunk", "Cannot flush instruction cache"
    End If
#Else
    Err.Raise 5, "ASF.Thunk", "Native execution requires Windows x64 Office and VBA7"
#End If
End Sub

Friend Function Execute(ByVal owner As ASF_VM, ByVal node As ASF_Map, ByVal scope As ASF_ScopeStack, ByVal expression As Boolean, ByRef result As Variant) As String
#If VBA7 And Win64 And Not Mac Then
    EnsureAvailable
    If programs Is Nothing Then Set programs = New Collection
    Dim key As String: key = CStr(ObjPtr(node)) & IIf(expression, "e", "s")
    Dim program As ASF_ThunkProgram
    On Error Resume Next
    Set program = programs(key)
    On Error GoTo 0
    If program Is Nothing Then
        Set program = New ASF_ThunkProgram
        program.Build node, expression
        programs.Add program, key
    End If
    On Error GoTo Failed
    Dim slots() As Variant, value As Variant, frame As NativeFrame
    ReDim slots(0 To program.Count)
    frame.code = program.CodeAddress
    frame.constants = program.ConstantsAddress
    frame.stack = VarPtr(slots(0))
    frame.count = program.Count
    frame.capacity = UBound(slots) + 1
    frame.constantCount = program.PoolCount
    frame.copy = copyAddress: frame.clear = clearAddress
    Dim status As Long, signal As String, op As String
    Do
        entries = entries + 1
        status = CLng(NativeInvoke(entry, frame, 0, 0, 0))
        If status = 0 Then Exit Do
        If status <> 1 Then Err.Raise 5, "ASF.Thunk", "Invalid native state or Variant HRESULT: " & CStr(frame.error)
        Select Case frame.request
            Case 2 ' host expression: root bypasses lowering; children can use it
                If frame.sp >= frame.capacity Then Err.Raise 5, "ASF.Thunk", "Stack overflow"
                CopyValue slots(frame.sp), owner.ThunkExpression(program.NodeAt(frame.argument), scope)
                frame.sp = frame.sp + 1
            Case 3 ' host statement: no replay of already executed operations
                signal = owner.ThunkStatement(program.NodeAt(frame.argument), scope)
                If Not program.IgnoreSignal(frame.pc) Then
                    Select Case signal
                        Case "BREAK": frame.signal = 1
                        Case "CONTINUE": frame.signal = 2
                        Case "RETURN": frame.signal = 3
                        Case "ERR": frame.signal = 4
                    End Select
                End If
            Case 4
                owner.ThunkAssign program.NodeAt(frame.argument), slots(frame.sp - 1), scope
                ClearValue slots(frame.sp - 1)
                frame.sp = frame.sp - 1
            Case 5
                scope.SetLocalValue "__return", slots(frame.sp - 1)
                ClearValue slots(frame.sp - 1)
                frame.sp = frame.sp - 1
                frame.signal = 3
            Case 9
                value = owner.ThunkTruthy(slots(frame.sp - 1))
                CopyValue slots(frame.sp - 1), value
            Case 10
                CopyValue value, owner.ThunkBinary(frame.argument, slots(frame.sp - 2), slots(frame.sp - 1))
                ClearValue slots(frame.sp - 1)
                CopyValue slots(frame.sp - 2), value
                ClearValue value
                frame.sp = frame.sp - 1
            Case 11
                op = program.NodeAt(frame.argument).GetValue("op")
                CopyValue value, owner.ThunkUnary(op, slots(frame.sp - 1))
                CopyValue slots(frame.sp - 1), value
                ClearValue value
            Case 15
                ' Let invokes an object's default member just where the AST VM
                ' does, before evaluation of the right operand/selected branch.
                value = slots(frame.sp - 1)
                CopyValue slots(frame.sp - 1), value
                ClearValue value
            Case Else
                Err.Raise 5, "ASF.Thunk", "Unknown host request"
        End Select
        frame.pc = frame.pc + 1
    Loop
    Select Case frame.signal
        Case 1: Execute = "BREAK"
        Case 2: Execute = "CONTINUE"
        Case 3: Execute = "RETURN"
        Case 4: Execute = "ERR"
    End Select
    If expression Then
        If frame.sp <> 1 Then Err.Raise 5, "ASF.Thunk", "Invalid expression stack"
        CopyValue result, slots(0)
    End If
    Exit Function
Failed:
    If Not expression Then
        If program.QuietFailure(frame.pc) Then
            Err.Clear
            Execute = "ERR"
            Exit Function
        End If
    End If
    Dim errorNumber As Long, errorSource As String, errorDescription As String
    errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
    Err.Raise errorNumber, errorSource, errorDescription
#Else
    EnsureAvailable
#End If
End Function

Private Sub CopyValue(ByRef destination As Variant, ByVal source As Variant)
    ' Hold an old object until the replacement is complete if its finalizer
    ' reenters ASF. Never leave a partially assigned Variant visible to reentry.
    Dim previous As Object
    If IsObject(destination) Then Set previous = destination
    If IsObject(source) Then
        Set destination = source
    Else
        If IsObject(destination) Then Set destination = Nothing
        destination = source
    End If
End Sub
Private Sub ClearValue(ByRef value As Variant)
    CopyValue value, Empty
End Sub
Private Sub Class_Terminate()
#If VBA7 And Win64 And Not Mac Then
    If entry <> 0 Then VirtualFree entry, 0, &H8000
#End If
End Sub
