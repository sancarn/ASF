Attribute VB_Name = "FastTests"
Option Explicit
Public BytecodeRuns As Long, ScopeRuns As Long, FallbackRuns As Long, NotRun As Long
Private checks As Long
Private currentSource As String
Public Sub RecordRoute(ByVal engine As ASF_VM)
    If Not ThunkProbe.FastDefault Then Exit Sub
    Dim log As Integer: log = FreeFile
    Open ThisWorkbook.Path & Chr(92) & "fast-routes.log" For Append As #log
    Print #log, engine.LastExecutionBackend & ":" & engine.FastFallbackReason
    Close #log
    Select Case engine.LastExecutionBackend
        Case "Array bytecode": BytecodeRuns = BytecodeRuns + 1
        Case "Scope bytecode": ScopeRuns = ScopeRuns + 1
        Case "VBA fallback": FallbackRuns = FallbackRuns + 1
        Case Else: NotRun = NotRun + 1
    End Select
End Sub
Private Sub Compare(ByVal source As String, Optional ByVal bytecode As Boolean = True)
    Dim pure As New ASF, fast As New ASF, p As Long, f As Long, i As Long
    currentSource = source
    fast.FastExecution = True
    p = pure.Compile(source): f = fast.Compile(source)
    For i = 1 To 2
        pure.Run p: fast.Run f
        checks = checks + 1
        If VarType(pure.OUTPUT_) <> VarType(fast.OUTPUT_) Then Err.Raise 5, , "type mismatch: " & source
        If Not IsNull(pure.OUTPUT_) Then
            If CStr(pure.OUTPUT_) <> CStr(fast.OUTPUT_) Then Err.Raise 5, , "value mismatch: " & source & " expected " & CStr(pure.OUTPUT_) & " got " & CStr(fast.OUTPUT_)
        End If
        If (bytecode And fast.LastExecutionBackend <> "Array bytecode") Or (Not bytecode And fast.LastExecutionBackend <> "Scope bytecode") Then Err.Raise 5, , "route mismatch: " & source & " " & fast.FastFallbackReason
        If Not bytecode And fast.FastScopeEntryCount = 0 Then Err.Raise 5, , "scope code was not entered"
        If fast.FastFallbackCount <> 0 Then Err.Raise 5, , "whole-program fallback"
        fast.FastOptimizations = (i = 2)
    Next i
End Sub
Public Function Run() As String
    On Error GoTo Failed
    checks = 0
    Compare "return (2+3)*7;"
    Compare "let s=0;for(let i=0;i<20;i+=1){if(i==3){continue;};if(i==10){break;};s+=i;};return s;"
    Compare "fun fib(n){if(n<=1){return n;};return fib(n-1)+fib(n-2);};return fib(14);"
    Compare "let x=5;fun f(){x+=2;return x;};return f()+x;"
    Compare "fun f(x){x+=2;return x;};let x=10;return f(3)+x;"
    Compare "fun f(n){if(n==0){return 0;};return 1+f(n-1);};return f(20);"
    Compare "return false && (1/0);"
    Compare "return true || (1/0);"
    Compare "return null ?? empty ?? 17;"
    Compare "return true ? 8 : (1/0);"
    Compare "return null == empty;"
    Compare "return 'a'+2;"
    Compare "return typeof([1,2,3]);"
    Compare "let x=0;try{x=1;x=1/0;}catch(e){x+=4;};return x;"
    Compare "let x=0;try{try{x=1/0;}catch(e){x=7;};x=1/0;}catch(e){x+=2;};return x;"
    Compare "fun f(){return 1/0;};return f() ?? 9;"
    Compare "return [1,2,3][2];"
    Compare "return [[1,2],[3,4]][2][1];"
    Compare "let o={a:2,b:{c:5}};return o.a+o.b.c;"
    Compare "let o={a:2};let s=0;for(let i=0;i<10;i+=1){s+=o.a;};return s;"
    Compare "let s=0;for(let x of [2,4,6]){s+=x;};return s;"
    Compare "let s=0;for(let x in [2,4,6]){s+=x;};return s;"
    Compare "let s='';for(let x of 'abc'){s+=x;};return s;"
    Compare "let s=0;for(let x of {a:3,b:7}){s+=x;};return s;"
    Compare "let s=0;for(let x of range(1,10)){if(x==2){continue;};if(x==5){break;};s+=x;};return s;"
    Compare "let x=2;switch(x){case 1 {return 10;}case 2 {return 20;}default {return 30;}};"
    Compare "let x=4;switch(x){case 1 {x=10;break;}default {x=30;}};return x;"
    Compare "return ' abc '.trim().toUpperCase()+' abc '.length;"
    Compare "return [1,2,3].length;"
    Compare "return isArray([1])+isNumeric('2');"
    Compare "return flatten([[1,2],[3,4]]).length;"
    ' Phases 2-5: assert scope bytecode entry as well as reference behavior.
    Compare "let x=1;let f=fun(){x+=1;return x;};return f();", False
    Compare "fun f(...r){return r.length;};return f(1,2,3);"
    Compare "let a=[1,2];return [...a,3].length;"
    Compare "let [a,b]=[2,5];return a+b;"
    Compare "class P{constructor(x){this.x=x;}get(){return this.x;}};let p=new P(4);return p.get();", False
    Compare "return @(1+2);", False
    Compare "let x=0;try{x=1;}catch(e){x=2;};return x;"
    Compare "let a=[1,2];a[2]=7;return a[2];"
    Compare "let a=[[1,2],[3,4]];a[2][1]=8;return a[2][1];"
    Compare "let a={x:2};a.x+=5;return a.x;"
    Compare "return `value ${2+3}`;"
    Compare "let a={x:2};let b={...a,y:5};return b.x+b.y;"
    Compare "let [a,...r]=[2,5,6];return a+r.length;"
    Compare "fun f(a,...r){return a+r.length;};return f(5,1,2);"
    Compare "fun f(a,b,...r){return a+b+r.length;};return f(5,7,1,2);"
    Compare "fun f(...r){return r.length;};return f();"
    Compare "fun f(...r){return r.length;};return f(empty);"
    Compare "return [...'abc'].length;"
    Compare "let n=0;fun f(){n+=1;return [n];};let a=[...f(),99];return n*10+a[1];"
    Compare "let n=0;fun f(){n+=1;return [n];};let a=clone(f());return n*10+a[1];"
    Compare "let x='a';x+=1;return x;"
    Compare "return $1foo ?? 9;"
    Compare "fun f(x){return x+1;};return f()+f(1,2);"
    Compare "let x=0;for(let i=1/0;false;i+=1){x=9;};return x;"
    Compare "fun f(){return 1;};let f=3;return f;", False
    Dim deep As New ASF, idx As Long
    deep.FastExecution = True
    idx = deep.Compile("fun f(n){if(n==0){return 0;};return 1+f(n-1);};return f(1000);")
    deep.Run idx
    If deep.OUTPUT_ <> 1000 Or deep.LastExecutionBackend <> "Array bytecode" Then Err.Raise 5, , "explicit frame growth"
    checks = checks + 1
    Compare "fun read(o){return o.a;};return read({a:1,b:2})+read({b:3,a:4});"
    Compare "let a={x:2};let b=a;b.x=5;return a.x;"
    Compare "let a=[1,2];let b=a;b[1]=5;return a[1];"
    Compare "return [] && true;"
    Compare "return `array ${[1,2]} null ${null}`;"
    Compare "let n=0;fun f(){n+=1;return [n];};let a=flatten(f());return n*10+a[1];"
    Compare "fun f(n){return n;};return f(...[1,2]);", False
    Compare "fun f(n){let s=0;for(let x of [1,2]){if(n>0){s+=f(n-1);};s+=x;};return s;};return f(2);"
    Compare "fun f(n){switch(n){case fcase(n){return 2;}default{return 3;}};};fun fcase(n){if(n>0){f(n-1);};return n;};return f(2);"
    Compare "return {classMethods:[],x:7}.x;"
    Compare "return null;"
    Compare "return empty;"
    Compare "return -3<<2;"
    Compare "return null+'x';"
    Compare "fun counter(){let n=0;return fun(){n+=1;return n;};};let f=counter();return f()+f();", False
    Compare "let a=[1,2,3];return a.map(fun(x){return x*2;}).reduce(fun(s,x){return s+x;},0);", False
    Compare "let x=0;let f=fun(){return x;};x=7;return f();", False
    Compare "let x=0;let f=fun(){x+=1;return x;};let g=fun(){return x;};f();return g();", False
    Compare "class P{constructor(x){this.x=x;}add(n){this.x+=n;return this.x;}};let p=new P(4);return p.add(3);", False
    Compare "let f=fun(n){if(n<=0){return 0;};return 1+f(n-1);};return f(20);", False
    Compare "let f=fun(n){if(n==0){return 1/0;};return f(n-1)??7;};return f(3);", False
    Compare "class P{constructor(x){this.x=x;}};let p=new P(4);return p.x;", False
    Compare "let f=fun(){let x=0;try{let a=[1,2];x=1/0;}catch(e){x=7;};return x;};return f();", False
    Compare "let f=fun(){try{for(let i=0;1/0;i+=1){break;};}catch(e){return 7;};return 3;};return f();", False
    Compare "let f=fun(){let x=0;try{break;x=7;}catch(e){x=9;};return x;};return f();", False
    Compare "class A{constructor(x){this.x=x;}f(){return this.x;}};class B extends A{constructor(x){super(x);}g(){return this.f()+2;}};let b=new B(5);return b.g();", False
    Compare "let x=2;let f=fun(){x+=3;return @(x*2);};return f()+x;", False
    Compare "let n=0;fun arg(){n+=1;return n;};let f=fun(x){return x;};return f(arg())*10+n;", False
    Compare "let f=fun(n){if(n==0){return 1/0;};f(0);return n;};return f(7);", False
    ExtraChecks
    Run = CStr(checks)
    Exit Function
Failed:
    Run = "ERROR " & Err.Number & " " & Err.Description & currentSource
End Function

Private Sub ExtraChecks()
    Dim scoped As New ASF, pScoped As Long, escaped As Variant
    scoped.FastExecution = True: scoped.AppAccess = True
    pScoped = scoped.Compile("let f=fun(n){if(n==0){return 0;};return 1+f(n-1);};return f(1000);")
    scoped.Run pScoped
    If scoped.OUTPUT_ <> 1000 Or scoped.FastScopeEntryCount < 1000 Then Err.Raise 5, , "dynamic closure trampoline growth"
    If scoped.FastClosureReleaseCount = 0 Then Err.Raise 5, , "closure arena was not released"
    checks = checks + 1
    pScoped = scoped.Compile("class P{f(n){if(n==0){return 0;};return 1+this.f(n-1);}};let p=new P();return p.f(500);")
    scoped.Run pScoped
    If scoped.OUTPUT_ <> 500 Then Err.Raise 5, , "bound-method trampoline growth"
    checks = checks + 1
    pScoped = scoped.Compile("fun counter(){let x=0;return fun(){x+=1;return x;};};return counter();")
    Set escaped = scoped.Run(pScoped)
    scoped.InjectVariable "escaped", escaped
    pScoped = scoped.Compile("return escaped()+escaped();")
    scoped.Run pScoped
    If scoped.OUTPUT_ <> 3 Then Err.Raise 5, , "escaped closure lifetime"
    scoped.Run pScoped
    If scoped.OUTPUT_ <> 7 Then Err.Raise 5, , "escaped closure shared state"
    checks = checks + 2
    Dim warm As New ASF, cacheSize As Long, repeatRun As Long, warmID As Long
    warm.FastExecution = True
    warmID = warm.Compile("class P{constructor(n){this.n=n;}get(){return this.n;}};let p=new P(7);let a=[1,2,3].map(fun(x){return x+1;});return p.get()+a.length;")
    warm.Run warmID: cacheSize = warm.FastScopeProgramCount
    For repeatRun = 1 To 50: warm.Run warmID: Next repeatRun
    If warm.OUTPUT_ <> 10 Or warm.FastScopeProgramCount <> cacheSize Then Err.Raise 5, , "scope cache retains transient values"
    checks = checks + 1
    Dim modules As New ASF, moduleID As Long, prototypeID As Long, entryCount As Double
    modules.FastExecution = True: modules.WorkingDir = ThisWorkbook.Path
    moduleID = modules.Compile("import Calculator from './calculator.vas';let c=Calculator();return c.add(2,3);")
    For repeatRun = 1 To 3
        entryCount = modules.FastScopeEntryCount: modules.Run moduleID
        If modules.OUTPUT_ <> 5 Or modules.FastScopeEntryCount <= entryCount Then Err.Raise 5, , "module exported closure bytecode"
    Next repeatRun
    checks = checks + 3
    modules.AppAccess = True: modules.OverrideCollMethods = True
    prototypeID = modules.Compile("prototype.COM.Worksheet Name(){return 'PATCHED';};return $1.Name;")
    modules.Run prototypeID, ThisWorkbook.Sheets(2)
    If modules.OUTPUT_ <> "PATCHED" Then Err.Raise 5, , "prototype registration"
    prototypeID = modules.Compile("return $1.Name;")
    modules.Run prototypeID, ThisWorkbook.Sheets(2)
    If modules.OUTPUT_ <> "PATCHED" Or modules.LastExecutionBackend <> "Scope bytecode" Then Err.Raise 5, , "persistent prototype dispatch"
    checks = checks + 2
    Dim engine As New ASF, pure As New ASF, idx As Long, old As Long, items As New Collection, arg As Variant
    engine.FastExecution = True
    old = engine.Compile("return 7;"): engine.Run old
    idx = engine.Compile("return 1/0;"): engine.Run idx
    If engine.OUTPUT_ <> 7 Or engine.FastRunCount <> 2 Then Err.Raise 5, , "root error changed output/replayed"
    checks = checks + 1
    ' Recompilation invalidates the cached call target/body of an earlier program.
    idx = engine.Compile("fun f(){return 1;};return f();"): engine.Run idx
    old = engine.Compile("fun g(){return 2;};return g();"): engine.Run old
    engine.Run idx
    If engine.OUTPUT_ <> 1 Then Err.Raise 5, , "cache invalidation"
    checks = checks + 1
    engine.AppAccess = True
    idx = engine.Compile("$1.Add(99);let f=fun(){return 1;};return $1.Count;")
    engine.Run idx, items
    If items.Count <> 1 Or engine.LastExecutionBackend <> "Scope bytecode" Then Err.Raise 5, , "fallback repeated side effect"
    checks = checks + 1
    idx = engine.Compile("return $1.Count+$1.Item(1);"): engine.Run idx, items
    If engine.OUTPUT_ <> 100 Or engine.LastExecutionBackend <> "Array bytecode" Then Err.Raise 5, , "COM boundary"
    checks = checks + 1
    engine.OverrideCollMethods = True: engine.Run idx, items
    pure.OverrideCollMethods = True: pure.AppAccess = True
    old = pure.Compile("return $1.Count+$1.Item(1);"): pure.Run old, items
    If CStr(engine.OUTPUT_) <> CStr(pure.OUTPUT_) Then Err.Raise 5, , "collection override"
    checks = checks + 1
    engine.NativeExecution = True: engine.Run idx, items
    If engine.FastExecution Or engine.NativeEntryCount = 0 Then Err.Raise 5, , "native selection"
    engine.FastExecution = True: engine.Run idx, items
    If engine.NativeExecution Then Err.Raise 5, , "bytecode selection"
    engine.FastExecution = False: engine.Run idx, items
    If engine.LastExecutionBackend <> "VBA AST" Then Err.Raise 5, , "AST selection"
    checks = checks + 1
    engine.FastExecution = True: engine.InjectVariable "x", 5
    idx = engine.Compile("return x;"): engine.Run idx
    If engine.OUTPUT_ <> 5 Or engine.LastExecutionBackend <> "Scope bytecode" Then Err.Raise 5, , "injected binding fallback"
    checks = checks + 1
End Sub
