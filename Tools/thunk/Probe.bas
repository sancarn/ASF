Attribute VB_Name = "ThunkProbe"
Option Explicit
Public NativeDefault As Boolean
Private checks As Long

Public Function ExistingSuite(ByVal native As Boolean) As String
    NativeDefault = native
    TestRunner.RunAllTests
    ExistingSuite = TestRunner.ThunkSuiteResult
    NativeDefault = False
End Function

Public Function Smoke() As String
    On Error GoTo Failed
    Dim engine As New ASF, idx As Long
    engine.NativeExecution = True
    idx = engine.Compile("return (2+3)*7;")
    engine.Run idx
    Smoke = CStr(engine.OUTPUT_) & ":" & CStr(engine.NativeEntryCount)
    Exit Function
Failed:
    Smoke = "ERROR " & Err.Number & " " & Err.Source & " " & Err.Description
End Function

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    checks = checks + 1
    If Not condition Then Err.Raise vbObjectError + 4000, "ThunkTests", message
End Sub

Public Function Regression() As String
    On Error GoTo Failed
    NativeDefault = False
    checks = 0
    Dim scripts As Variant, script As Variant, pure As ASF, native As ASF, p As Long, n As Long, repeat As Long
    scripts = Array( _
        "return (2+3)*7;", _
        "let s=0; for(let i=0;i<20;i+=1){if(i==3){continue;};if(i==10){break;};s+=i;};return s;", _
        "let i=0; let s=0; while(i<20){i+=1;if(i<3){continue;};if(i==9){break;};s+=i;};return s;", _
        "let i=0; while(i<4){let j=0;while(j<3){j+=1;if(j==2){continue;};};i+=1;};return i;", _
        "return false && (1/0);", "return true || (1/0);", _
        "return null ?? empty ?? 17;", "return true ? 8 : (1/0);", _
        "fun fib(n){if(n<2){return n;};return fib(n-1)+fib(n-2);};return fib(8);", _
        "return [1,2,3,4].map(fun(x){return x*2;}).reduce(fun(a,x){return a+x;},0);", _
        "let x=0;try{ x=1; x=1/0;}catch(e){x+=4;};return x;", _
        "return 'a'+2;", "return null == empty;", "return 7%3;", _
        "return 2^8;", "return 3<<2;", "return typeof(12);", _
        "let x=3;if(x==1){return 1;}elseif(x==2){return 2;}else{return 3;};")
    Dim answers As Variant, scriptIndex As Long
    answers = Array("35", "42", "33", "4", "False", "True", "17", "8", "21", "20", "5", "a2", "True", "1", "256", "12", "number", "3")
    For Each script In scripts
        Set pure = New ASF: Set native = New ASF
        native.NativeExecution = True
        p = pure.Compile(CStr(script)): n = native.Compile(CStr(script))
        For repeat = 1 To 3
            pure.Run p: native.Run n
            Check CStr(pure.OUTPUT_) = CStr(answers(scriptIndex)), "expected answer: " & CStr(script) & " actual=" & CStr(pure.OUTPUT_)
            Check VarType(pure.OUTPUT_) = VarType(native.OUTPUT_), "subtype: " & CStr(script)
            Check CStr(pure.OUTPUT_) = CStr(native.OUTPUT_), "value: " & CStr(script)
        Next repeat
        Check native.NativeEntryCount > 0, "native routing"
        Check pure.NativeEntryCount = 0, "pure routing"
        native.ClearNativeCache
        Check native.NativeEntryCount = 0, "cache clear"
        native.Run n
        Check CStr(pure.OUTPUT_) = CStr(native.OUTPUT_), "after cache clear"
        native.NativeExecution = False
        native.Run n
        Check CStr(pure.OUTPUT_) = CStr(native.OUTPUT_), "toggle back to VBA"
        scriptIndex = scriptIndex + 1
    Next script
    VariantCases
    ReentryCases
    Regression = CStr(checks)
    Exit Function
Failed:
    Regression = "ERROR " & Err.Number & " " & Err.Description
End Function

Private Function Describe(ByVal value As Variant) As String
    If IsObject(value) Then
        Describe = "object:" & TypeName(value)
    ElseIf IsArray(value) Then
        Dim item As Variant
        Describe = "array:"
        For Each item In value
            Describe = Describe & Describe(item) & ";"
        Next item
    ElseIf IsNull(value) Then
        Describe = "Null"
    ElseIf IsEmpty(value) Then
        Describe = "Empty"
    Else
        Describe = VarType(value) & ":" & CStr(value)
    End If
End Function
Private Function Outcome(ByVal engine As ASF, ByVal idx As Long, ByVal a As Variant, ByVal b As Variant) As String
    On Error GoTo Failed
    engine.Run idx, a, b
    Outcome = Describe(engine.OUTPUT_)
    Dim globals As ASF_Globals: Set globals = engine.GetGlobals
    Dim line As Variant
    For Each line In globals.gRuntimeLog
        Outcome = Outcome & "|" & CStr(line)
    Next line
    Exit Function
Failed:
    Outcome = "ERROR:" & Err.Number & ":" & Err.Source & ":" & Err.Description
End Function
Private Sub VariantCases()
    Dim inputs As Variant, expressions As Variant, expression As Variant, a As Variant, b As Variant
    inputs = Array(Empty, Null, True, False, CByte(2), CInt(-3), CLng(5), CSng(1.25), CDbl(2.5), _
                   CCur(3.2), CDec(4.1), CDate(45000), "2", "false", "abc", CVErr(2042), Array(1, 2))
    expressions = Array("return $1+$2;", "return $1-$2;", "return $1*$2;", "return $1/$2;", _
                        "return $1<$2;", "return $1==$2;", "return $1!=$2;", "return $1%$2;", _
                        "return $1 && $2;", "return $1 || $2;", "return $1 ?? $2;", _
                        "return $1 ? $2 : 3;", "return !$1;", "return -$1;", "return typeof $1;", _
                        "for(let i=0;1/0;i+=1){break;};return 7;", _
                        "if(true){for(let i=0;1/0;i+=1){break;};};return 7;")
    Dim pure As ASF, native As ASF, p As Long, n As Long, i As Long
    Dim expected As String, actual As String
    For Each expression In expressions
        Set pure = New ASF: Set native = New ASF
        pure.AppAccess = True: native.AppAccess = True
        pure.verbose = True: native.verbose = True
        native.NativeExecution = True
        p = pure.Compile(CStr(expression)): n = native.Compile(CStr(expression))
        For i = 0 To UBound(inputs)
            If IsObject(a) Then Set a = Nothing
            a = inputs(i)
            b = inputs((i + 3) Mod (UBound(inputs) + 1))
            Set pure.GetGlobals.gRuntimeLog = New Collection
            Set native.GetGlobals.gRuntimeLog = New Collection
            expected = Outcome(pure, p, a, b)
            actual = Outcome(native, n, a, b)
            Check expected = actual, CStr(expression) & " input " & i & " expected " & expected & " actual " & actual
        Next i
    Next expression
End Sub
Private Sub ReentryCases()
    Dim native As Long, repeat As Long, engine As ASF, probe As ThunkReentryProbe, idx As Long
    Dim baselineCalls As Long
    For native = 0 To 1
        Set engine = New ASF
        engine.NativeExecution = CBool(native)
        engine.AppAccess = True
        Set probe = New ThunkReentryProbe
        Set probe.engine = engine
        probe.inner = engine.Compile("return 5;")
        idx = engine.Compile("return $1 + $1;")
        For repeat = 1 To 30
            engine.Run idx, probe
            Check CDbl(engine.OUTPUT_) = 10, "default member reentry result"
        Next repeat
        If native = 0 Then baselineCalls = probe.calls
        Check probe.calls = baselineCalls, "default member reentry count: " & probe.calls & " vs " & baselineCalls
        idx = engine.Compile("let a=$1.Make();a=0;return 7;")
        For repeat = 1 To 30
            engine.Run idx, probe
            Check CDbl(engine.OUTPUT_) = 7, "finalizer reentry result"
            Check probe.releases = repeat, "finalizer release count"
        Next repeat
        Set probe.engine = Nothing
        Set probe = Nothing: Set engine = Nothing
    Next native
End Sub
