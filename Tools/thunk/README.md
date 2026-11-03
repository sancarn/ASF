# Experimental ASF native VM

ASF's compiler builds an AST of `ASF_Map` nodes and `Collection` blocks. Its
existing VM recursively interprets that AST; it is **not a bytecode VM**.
This backend adds an AST-to-instruction lowering pass and an embedded,
relocation-free C interpreter for Windows x64 Office. The parser and public
Compile/Run contract stay the same.

```vb
Dim engine As New ASF
engine.NativeExecution = True   ' opt in before running; default is False
Dim program As Long
program = engine.Compile("let s=0; for(let i=0;i<100;i+=1){s+=i;};return s;")
Debug.Print engine.Run(program) ' 4950
Debug.Print engine.NativeEntryCount ' confirms calls actually entered C
engine.NativeExecution = False ' select the VBA AST evaluator again
```

Import **all fourteen classes in `src/`**, including `ASF_Thunk.cls` and
`ASF_ThunkProgram.cls`. These two additional imports are required when upgrading
the ASF/ASF_VM classes, even when native execution is disabled. There is no
external runtime DLL to distribute. `oleaut32.dll`, `kernel32.dll` and
`user32.dll` are system dependencies of the native backend.

Native execution requires Windows x64 Office with VBA7. Enabling it on other
Office hosts raises error 5 before execution. VBA execution remains the default;
the LibreOffice builder omits the Windows backend and retains its AST engine.
The native API is experimental, not a claim of universal semantic parity or speed.

## Execution and ownership

`ASF_ThunkProgram` lowers each requested statement/expression once. Instructions
are four 32-bit fields (opcode, argument, break target, continue target).
Constants are real Automation Variants. C owns the program counter, operand
stack position, conditionals, short-circuit jumps, classic for/while loops and
break/continue/return signals. Arithmetic/comparisons on two finite Doubles use
native operations; other subtypes, coercions and ASF-specific operators use the
same VBA operations as the reference evaluator.

The interpreter yields to VBA for variable access, assignments, complex
expressions and language services. Objects, arrays, closures, calls, builtins,
COM, regex, classes, modules, for-in/of, switch and try/catch retain the existing
host implementations. Their child expressions/statements can enter native code.
Thus this is a hybrid VM, not a standalone replacement for the whole ASF runtime.
No program is restarted in VBA after native execution has begun; side effects
are never replayed as a fallback strategy.

Every invocation owns its frame, Variant stack and temporary values. Cached
programs retain their AST roots, preventing reuse of pointer cache keys. Active
calls pin the backend and program even when a callback clears the instance cache.
Executable bytes and Automation helper addresses are cached per VM. Code is
allocated read/write, changed to execute/read, then its instruction cache is
flushed. It is never allocated read/write/execute. Normal class termination frees
it; VBE Reset/End bypasses termination and can leak an allocation.

The public `GetGlobals` API exposes mutable AST objects. After changing these
objects directly, call `engine.ClearNativeCache` before the next run. Switching
to a different Globals object also invalidates the cache. Clearing it resets
`NativeEntryCount`; toggling native mode alone retains compiled programs.
Do not mutate an AST that is actively executing.

## Lessons carried over from stdLambda

- Lower the source representation explicitly: ASF starts with AST nodes, unlike
  stdLambda's existing operation stream.
- Cache code and helper addresses; count cold setup separately from warm runs.
- Keep mutable execution state local to support callbacks and finalizer reentry.
- Use real Variants, reference ownership and ASF's coercions rather than replacing
  the language with a numeric-only approximation.
- Yield for host behavior instead of rerunning partially executed source.
- Check generated code for relocations and verify layouts at build time.
- Judge speed from measured workloads, not from the choice of implementation language.

## Build and validate

From the repository root, with MinGW-w64 GCC and Python on Windows:

```powershell
python Tools/thunk/build.py --toolchain C:/Ruby31-x64/msys64/ucrt64/bin
python Tools/thunk/build.py --toolchain C:/Ruby31-x64/msys64/ucrt64/bin --check
python Tools/thunk/run_excel.py --output Tools/thunk/results/validation
python Tools/thunk/run_excel.py --baseline --output Tools/thunk/results/baseline
python libreoffice/tools/build.py
```

The C builder runs control, numeric, ownership and bounds tests, asserts the
24-byte x64 VARIANT, 16-byte instruction and 80-byte frame layouts, rejects all
object relocations, checks the entry symbol is at offset zero and embeds `.text`
between generated markers in `ASF_Thunk.cls`. `--check` rejects stale payloads.
The checked-in payload is reproducible with GCC 11.2.0 from the listed toolchain;
other compiler versions may generate different correct bytes and require rebuilding.

Excel testing needs `pywin32` and `psutil`. The supervisor follows the requested
stdVBA `AGENTS.md` rules: disposable hidden instance, CRLF checks before every
import, owned-window dialog detection/dismissal, bounded execution and exact
restoration of AccessVBOM and VBE error-trapping registry values/types/presence.
Do not run Office supervisors concurrently: those registry preferences are shared.
The harness currently recognizes installed Office 16/VBA 7.1. It refuses to guess
the debugger key for other versions. A dismissed dialog always fails the run.

The test workbook adds only harness hooks: a default-engine selector and test
summary accessor. It imports the repository's original TestRunner and `.vas`
fixtures. `--baseline` tests commit `527f172` with only the private ByVal
compilation repair described below. It does not include native execution.

## Validation and limitations

Excel 16 x64 validation passes **723 additional assertions** for expected
answers, repeated calls, cache reset, mode switching, Variant subtypes, Null,
Empty, strings, errors, arrays, coercion/error logs, nested loops, short circuits,
recursion, callbacks and same-instance default-member/finalizer reentry.
The original suite has **254/255 passing in both engines**, with identical
outcomes. The pre-change source reproduces the remaining failure:
`host_sees_array_mutation` returns `8` instead of `42`. The harness allows only
that exact known failure and rejects additional failures or engine disagreement.
This is recorded as a baseline limitation, not a passing test. Checked-in
[evidence](validation/excel.json) includes source hashes and [baseline results](validation/baseline.json).

The existing named-function path also passed a Variant array element to a ByRef
Collection parameter, which Excel rejected at compile time. `InvokeClosure` now
takes that Collection reference ByVal. The referenced Collection and execution
semantics are unchanged; the private parameter was never reassigned.

Only Excel Windows x64 execution was validated. The LibreOffice extension builds,
but no new LibreOffice runtime tests or x86/Mac Office tests were run. Exhaustive
locale, COM object, resource exhaustion and finalizer-lifetime parity is not
claimed. Stack/copy lifetimes differ internally, so custom finalizers outside the
covered cases remain a boundary of this experiment. The original VM's error
containment and stale OUTPUT_ behavior are preserved, including its known defects.
Native code is not a sandbox for untrusted input; the reference host services
still determine access to Office and external objects.

Benchmark tooling/results live on the companion `benchmark/thunk-vs-vba` branch.
They compare this branch's reference VBA mode with its native mode, not stdLambda
or a different ASF release. See that branch's `Tools/thunk/benchmark/README.md`.
