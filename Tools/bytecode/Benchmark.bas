Attribute VB_Name = "ThunkBenchmark"
Option Explicit
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
Private source As String, expected As String, args As Variant
Private argument1 As Variant, argument2 As Variant

Private Function Seconds() As Double
    Static frequency As Currency
    Dim counter As Currency
    If frequency = 0 Then QueryPerformanceFrequency frequency
    QueryPerformanceCounter counter
    Seconds = counter / frequency
End Function
Private Sub Configure(ByVal scenario As Long)
    Dim items As Collection, numbers() As Variant, i As Long
    If IsObject(argument1) Then Set argument1 = Nothing
    If IsObject(argument2) Then Set argument2 = Nothing
    argument1 = Empty: argument2 = Empty
    Select Case scenario
        Case 0
            source = "return ($1 + $2) * 3 - $1 / 2;"
            args = Array(7, 5): expected = "32.5"
        Case 1
            source = "return $1 % 3 == 0 && $1 > 10;"
            args = Array(12): expected = "True"
        Case 2
            source = "return $1.trim().toUpperCase() + $1.length;"
            args = Array("  abc  "): expected = "ABC7"
        Case 3
            source = "return $1.count + $1.item(2);"
            Set items = New Collection
            items.Add 10: items.Add 20: items.Add 30
            args = Array(items): expected = "23"
        Case 4
            source = "let s=0; for(let i=1;i<=2000;i+=1){s=s+i*2-1;};return s;"
            args = Array(): expected = "4000000"
        Case 5
            source = "let i=0;let n=0;while(i<1000){if(i%3==0){n=n+2;}else{n=n-1;};i+=1;};return n;"
            args = Array(): expected = "2"
        Case 6
            source = "fun fib(n){if(n<=1){return n;};return fib(n-1)+fib(n-2);};return fib($1);"
            args = Array(14): expected = "377"
        Case 7
            source = "let s=0;$1.forEach(fun(x){s=s+x*2+1;});return s;"
            ReDim numbers(1 To 300)
            For i = 1 To 300: numbers(i) = i: Next i
            args = Array(numbers): expected = "90600"
        Case 8
            source = "let s = 0; for (let i = 1, i <= 2000, i += 1) { s = s + i * 2 - 1; }; return(s);"
            args = Array(): expected = "4000000"
        Case 9
            source = "fun fib(n) { if (n <= 1) { return n; }; return fib(n - 1) + fib(n - 2); }; return(fib(14));"
            args = Array(): expected = "377"
        Case 10
            source = "let s = ''; for (let i = 1, i <= 500, i += 1) { s = s + 'x' + i; }; return(s.length);"
            args = Array(): expected = "1892"
        Case 11
            source = "let a = []; for (let i = 1, i <= 300, i += 1) { a.push(i); }; let t = a.map(fun(x) { return x * 2; }).filter(fun(x) { return x % 3 == 0; }).reduce(fun(acc, v) { return acc + v; }, 0); return(t);"
            args = Array(): expected = "30300"
        Case 12
            source = "let o = { a: 0, b: 1, c: { d: 2 } }; for (let i = 1, i <= 1000, i += 1) { o.a = o.a + o.b + o.c.d; }; return(o.a);"
            args = Array(): expected = "3000"
        Case 13
            source = "fun counter() { let c = 0; return fun() { c = c + 1; return c; }; }; let f = counter(); let r = 0; for (let i = 1, i <= 500, i += 1) { r = f(); }; return(r);"
            args = Array(): expected = "500"
        Case 14
            source = "let i = 0; let n = 0; while (i < 1000) { if (i % 3 == 0) { n = n + 2; } else { n = n - 1; }; i += 1; }; return(n);"
            args = Array(): expected = "2"
        Case 15
            source = "let out = 0; for (let i = 1, i <= 300, i += 1) { let s = `item ${i}`; out = out + s.toUpperCase().length; }; return(out);"
            args = Array(): expected = "2292"
        Case 16
            source = "let a = [0, 0, 0, 0, 0, 0]; for (let i = 1, i <= 300, i += 1) { a[1 + i % 6] = a[1 + i % 6] + i; }; return(a[1] + a[2] + a[3] + a[4] + a[5] + a[6]);"
            args = Array(): expected = "45150"
        Case 17
            source = "let m = [[0, 0, 0], [0, 0, 0], [0, 0, 0]]; for (let i = 1, i <= 300, i += 1) { m[1 + i % 3][1 + i % 2] = m[1 + i % 3][1 + i % 2] + i; }; return(m[1][1] + m[2][2] + m[3][1]);"
            args = Array(): expected = "22500"
        Case 18
            source = "class P { constructor(x) { this.x = x; } add(d) { this.x = this.x + d; return this.x; } }; let p = new P(0); for (let i = 1, i <= 300, i += 1) { p.add(i); }; return(p.x);"
            args = Array(): expected = "45150"
        Case 19
            source = "let records=[{qty:3,cost:12},{qty:7,cost:4},{qty:2,cost:25}];let total=0;for(let r of records){total+=r.qty*r.cost;};return total;"
            args = Array(): expected = "114"
        Case 20
            source = "let records=[{hours:8,rate:20,active:true},{hours:6,rate:25,active:false},{hours:10,rate:30,active:true}];return records.filter(fun(r){return r.active;}).reduce(fun(s,r){return s+r.hours*r.rate;},0);"
            args = Array(): expected = "460"
        Case 21
            source = "let s=0;for(let i=0;i<100;i+=1){try{let x=10/(i%5);s+=x;}catch(e){s+=1;};};return s;"
            args = Array(): expected = "436.666666666667"
        Case Else
            Err.Raise 5, "ThunkBenchmark", "Unknown scenario"
    End Select
    ' Materialize array elements before passing them to ASF's ParamArray. The
    ' reference engine misreads a ByRef array-element wrapper on this Office host.
    If UBound(args) >= 0 Then
        If IsObject(args(0)) Then Set argument1 = args(0) Else argument1 = args(0)
    End If
    If UBound(args) >= 1 Then
        If IsObject(args(1)) Then Set argument2 = args(1) Else argument2 = args(1)
    End If
End Sub
Private Function NewEngine(ByVal backend As Long, ByRef idx As Long) As ASF
    Dim engine As New ASF
    engine.NativeExecution = (backend = 1)
    engine.FastExecution = (backend >= 2)
    engine.FastOptimizations = (backend <> 3)
    engine.AppAccess = True
    idx = engine.Compile(source)
    Set NewEngine = engine
End Function
Private Sub RunOnce(ByVal engine As ASF, ByVal idx As Long)
    Select Case UBound(args)
        Case -1: engine.Run idx
        Case 0: engine.Run idx, argument1
        Case 1: engine.Run idx, argument1, argument2
    End Select
End Sub
Private Sub Validate(ByVal engine As ASF, ByVal backend As Long)
    If CStr(engine.OUTPUT_) <> expected Then Err.Raise vbObjectError + 4001, "ThunkBenchmark", "Wrong answer: " & CStr(engine.OUTPUT_) & "; expected " & expected
    If backend = 1 Then
        If engine.NativeEntryCount <= 0 Then Err.Raise 5, "ThunkBenchmark", "Native route was not entered"
    Else
        If engine.NativeEntryCount <> 0 Then Err.Raise 5, "ThunkBenchmark", "VBA route entered native code"
    End If
    If backend >= 2 Then
        If engine.FastFallbackCount <> 0 Then Err.Raise 5, , "Bytecode benchmark used whole-program fallback"
        If engine.LastExecutionBackend = "Scope bytecode" Then
            If engine.FastScopeEntryCount = 0 Then Err.Raise 5, , "Scope bytecode not entered"
        ElseIf engine.LastExecutionBackend <> "Array bytecode" Then
            Err.Raise 5, , "Unexpected bytecode route"
        End If
    End If
End Sub

' 0: new engine + configure backend + compile AST (no Run in the timer)
' 1: new engine + configure backend + compile AST + first Run
' 2: repeated Run after one validation/warm-up call (setup outside the timer)
Public Function Measure(ByVal backend As Long, ByVal scenario As Long, ByVal phase As Long, ByVal count As Long) As String
    On Error GoTo Failed
    If count < 1 Or phase < 0 Or phase > 2 Then Err.Raise 5
    ThunkProbe.NativeDefault = False: ThunkProbe.FastDefault = False
    Configure scenario
    Dim engine As ASF, idx As Long, i As Long, started As Double, elapsed As Double
    Set engine = NewEngine(backend, idx)
    RunOnce engine, idx
    Validate engine, backend
    started = Seconds()
    Select Case phase
        Case 0
            For i = 1 To count
                Set engine = NewEngine(backend, idx)
            Next i
        Case 1
            For i = 1 To count
                Set engine = NewEngine(backend, idx)
                RunOnce engine, idx
            Next i
        Case 2
            For i = 1 To count
                RunOnce engine, idx
            Next i
    End Select
    elapsed = Seconds() - started
    If phase = 0 Then RunOnce engine, idx
    Validate engine, backend
    Measure = Trim$(Str$(elapsed)) & "|" & engine.LastExecutionBackend & "|" & engine.FastFallbackReason
    Exit Function
Failed:
    Measure = "ERROR " & Err.Number & " " & Err.Source & " " & Err.Description
End Function
