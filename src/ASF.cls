VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private GLOBALS_ As ASF_Globals
Private Compiler_ As ASF_Compiler
Private VM_ As ASF_VM
Private OUTPUT__ As Variant
Private EVALUATOR_ As VBAexpressions

Public Property Get verbose() As Boolean
    verbose = VM_.verbose
End Property
Public Property Let verbose(aValue As Boolean)
    VM_.verbose = aValue
End Property
Public Property Get EnableCallTrace() As Boolean
    EnableCallTrace = VM_.EnableCallTrace
End Property
Public Property Let EnableCallTrace(aValue As Boolean)
    VM_.EnableCallTrace = aValue
End Property
Public Property Get GetCallStackTrace() As String
    GetCallStackTrace = VM_.GetCallStackTrace
End Property
Public Property Get GetCallStack() As Collection
    Set GetCallStack = VM_.GetCallStack
End Property
Public Property Let WorkingDir(baseDir As String)
    GLOBALS_.CURRENT_MODULE_PATH = baseDir
End Property
Public Property Get WorkingDir() As String
    WorkingDir = GLOBALS_.CURRENT_MODULE_PATH
End Property
Public Property Get Evaluator() As VBAexpressions
    Set Evaluator = EVALUATOR_
End Property
Public Property Get CastInjectedVars() As Boolean
    CastInjectedVars = VM_.CastInjectedVars
End Property
Public Property Let CastInjectedVars(aValue As Boolean)
    VM_.CastInjectedVars = aValue
End Property
Public Property Get AppAccess() As Boolean
    AppAccess = VM_.AppAccess
End Property
Public Property Let AppAccess(aValue As Boolean)
    VM_.AppAccess = aValue
End Property
Public Property Get OverrideCollMethods() As Boolean
    OverrideCollMethods = VM_.OverrideCollMethods
End Property
Public Property Let OverrideCollMethods(aValue As Boolean)
    VM_.OverrideCollMethods = aValue
End Property
Private Sub Class_Initialize()
    INIT_
End Sub
Public Function Compile(src As String, Optional progName As String = "@anon") As Long
    If GLOBALS_ Is Nothing Then
        INIT_
    End If
    With Compiler_
        .SetGlobals GLOBALS_
        Compile = .CompileProgram(src, progName)
' BEGIN FAST BYTECODE BACKEND
        VM_.ClearFastCache
' END FAST BYTECODE BACKEND
    End With
End Function
Public Function CreateEvaluator(expression As String) As VBAexpressions
    Set EVALUATOR_ = New VBAexpressions
    EVALUATOR_.Create expression
    Set CreateEvaluator = EVALUATOR_
End Function
Public Function GetGlobals()
    Set GetGlobals = GLOBALS_
End Function
Public Function SetGlobals(aGlobals As ASF_Globals)
    Set GLOBALS_ = aGlobals
End Function
Public Sub InjectVariable(varName As String, VarValue As Variant)
    VM_.InjectVariable varName, VarValue
End Sub
Private Sub INIT_()
    Set GLOBALS_ = New ASF_Globals
    GLOBALS_.ASF_InitGlobals
    Set Compiler_ = New ASF_Compiler
    Set VM_ = New ASF_VM
' BEGIN FAST BYTECODE BACKEND
    VM_.FastOptimizations = True
' END FAST BYTECODE BACKEND
    VM_.verbose = False
    VM_.EnableCallTrace = False
    VM_.AppAccess = False
    VM_.CastInjectedVars = True
    VM_.OverrideCollMethods = False
    Set VM_.Globals = GLOBALS_
End Sub
Public Property Get OUTPUT_() As Variant
    vAssignment OUTPUT_, OUTPUT__
End Property
Public Function Execute(scriptPath As String, ParamArray params() As Variant) As Variant
    If Not GLOBALS_ Is Nothing Then
        With VM_
            Set .Globals = GLOBALS_
            If Not IsMissing(params) Then
                Dim params_ As Variant: params_ = params
                .RunProgramByIndex Compile(ReadTextFile(scriptPath)), params_
            Else
                .RunProgramByIndex Compile(ReadTextFile(scriptPath))
            End If
            vAssignment OUTPUT__, .OUTPUT_
            vAssignment Execute, .OUTPUT_
                If .verbose Then
                    Debug.Print "=== Runtime Log ==="
                    Dim elm As Variant
                    For Each elm In GLOBALS_.gRuntimeLog
                        Debug.Print elm
                    Next elm
                End If
        End With
    End If
End Function
Public Function Eval(Optional strValues As String) As String
    With EVALUATOR_
        Eval = .Eval(strValues)
    End With
End Function
Public Function Run(idx As Long, ParamArray params() As Variant) As Variant
    If Not GLOBALS_ Is Nothing Then
        With VM_
            Set .Globals = GLOBALS_
            If Not IsMissing(params) Then
                Dim params_ As Variant: params_ = params
                .RunProgramByIndex idx, params_
            Else
                .RunProgramByIndex idx
            End If
            vAssignment OUTPUT__, .OUTPUT_
            vAssignment Run, .OUTPUT_
                If .verbose Then
                    Debug.Print "=== Runtime Log ==="
                    Dim elm As Variant
                    For Each elm In GLOBALS_.gRuntimeLog
                        Debug.Print elm
                    Next elm
                End If
        End With
    End If
End Function
Private Sub vAssignment(ByRef var As Variant, ByRef vValue As Variant)
    If IsObject(vValue) Then
        Set var = vValue
    Else
        var = vValue
    End If
End Sub
Public Sub ClearCallStack()
    VM_.ClearCallStack
End Sub
Public Function ReadTextFile(filePath As String) As String
    ReadTextFile = VM_.ReadTextFile(filePath)
End Function
Public Sub ClearModuleCache()
    If Not GLOBALS_ Is Nothing Then
        Set GLOBALS_.gModuleRegistry = New ASF_Map
        Set GLOBALS_.gModuleExports = New ASF_Map
        Set GLOBALS_.gLoadingModules = New Collection
        GLOBALS_.CURRENT_MODULE_PATH = vbNullString
    End If
End Sub
Public Function VarToASF(VarValue) As Variant
    VarToASF = VM_.VarToASF(VarValue)
End Function

' BEGIN WINDOWS NATIVE BACKEND
' Opt-in Windows x64 C thunk backend. False keeps the reference VBA AST VM.
' Set before Run/Execute; parsing and compilation remain in VBA.
Public Property Let NativeExecution(ByVal enabled As Boolean)
    VM_.NativeExecution = enabled
End Property
Public Property Get NativeExecution() As Boolean
    NativeExecution = VM_.NativeExecution
End Property
' Diagnostic count of native entries; useful to verify benchmark routing.
Public Property Get NativeEntryCount() As Double
    NativeEntryCount = VM_.NativeEntryCount
End Property
' Required after externally mutating AST nodes obtained through GetGlobals.
Public Sub ClearNativeCache()
    VM_.ClearNativeCache
End Sub
' END WINDOWS NATIVE BACKEND

' BEGIN FAST BYTECODE BACKEND
' Experimental bytecode execution: compact slots or shared-scope instructions.
' Complex services re-enter bytecode for their script bodies; no AST replay.
Public Property Let FastExecution(ByVal enabled As Boolean)
    VM_.FastExecution = enabled
End Property
Public Property Get FastExecution() As Boolean
    FastExecution = VM_.FastExecution
End Property
Public Property Let FastOptimizations(ByVal enabled As Boolean)
    VM_.FastOptimizations = enabled
End Property
Public Property Get FastOptimizations() As Boolean
    FastOptimizations = VM_.FastOptimizations
End Property
Public Property Get LastExecutionBackend() As String
    LastExecutionBackend = VM_.LastExecutionBackend
End Property
Public Property Get FastFallbackReason() As String
    FastFallbackReason = VM_.FastFallbackReason
End Property
Public Property Get FastRunCount() As Long
    FastRunCount = VM_.FastRunCount
End Property
Public Property Get FastFallbackCount() As Long
    FastFallbackCount = VM_.FastFallbackCount
End Property
' Call after externally mutating GetGlobals' AST/function tables. Compile invalidates automatically.
Public Sub ClearFastCache()
    VM_.ClearFastCache
End Sub
' END FAST BYTECODE BACKEND

' BEGIN FAST BYTECODE BACKEND
Public Property Get FastScopeEntryCount() As Double
    FastScopeEntryCount = VM_.FastScopeEntryCount
End Property
Public Property Get FastHostServiceCount() As Double
    FastHostServiceCount = VM_.FastHostServiceCount
End Property
' END FAST BYTECODE BACKEND

' BEGIN FAST BYTECODE BACKEND
Public Property Get FastClosureReleaseCount() As Double
    FastClosureReleaseCount = VM_.FastClosureReleaseCount
End Property
' END FAST BYTECODE BACKEND

' BEGIN FAST BYTECODE BACKEND
Public Property Get FastScopeProgramCount() As Long
    FastScopeProgramCount = VM_.FastScopeProgramCount
End Property
' END FAST BYTECODE BACKEND
