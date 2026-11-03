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
Private Function NewEngine(ByVal native As Boolean, ByRef idx As Long) As ASF
    Dim engine As New ASF
    engine.NativeExecution = native
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
Private Sub Validate(ByVal engine As ASF, ByVal native As Boolean)
    If CStr(engine.OUTPUT_) <> expected Then Err.Raise vbObjectError + 4001, "ThunkBenchmark", "Wrong answer: " & CStr(engine.OUTPUT_) & "; expected " & expected
    If native Then
        If engine.NativeEntryCount <= 0 Then Err.Raise 5, "ThunkBenchmark", "Native route was not entered"
    Else
        If engine.NativeEntryCount <> 0 Then Err.Raise 5, "ThunkBenchmark", "VBA route entered native code"
    End If
End Sub

' 0: new engine + configure backend + compile AST (no Run in the timer)
' 1: new engine + configure backend + compile AST + first Run
' 2: repeated Run after one validation/warm-up call (setup outside the timer)
Public Function Measure(ByVal native As Boolean, ByVal scenario As Long, ByVal phase As Long, ByVal count As Long) As String
    On Error GoTo Failed
    If count < 1 Or phase < 0 Or phase > 2 Then Err.Raise 5
    ThunkProbe.NativeDefault = False
    Configure scenario
    Dim engine As ASF, idx As Long, i As Long, started As Double, elapsed As Double
    Set engine = NewEngine(native, idx)
    RunOnce engine, idx
    Validate engine, native
    started = Seconds()
    Select Case phase
        Case 0
            For i = 1 To count
                Set engine = NewEngine(native, idx)
            Next i
        Case 1
            For i = 1 To count
                Set engine = NewEngine(native, idx)
                RunOnce engine, idx
            Next i
        Case 2
            For i = 1 To count
                RunOnce engine, idx
            Next i
    End Select
    elapsed = Seconds() - started
    If phase = 0 Then RunOnce engine, idx
    Validate engine, native
    Measure = Trim$(Str$(elapsed))
    Exit Function
Failed:
    Measure = "ERROR " & Err.Number & " " & Err.Source & " " & Err.Description
End Function
