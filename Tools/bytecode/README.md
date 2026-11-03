# Garcia bytecode backend experiment

This branch explores the five-phase proposal in
[WS Garcia's PR comment](https://github.com/sancarn/ASF/pull/1#issuecomment-6600185041).
It is **an incomplete five-phase implementation**, not a replacement for the
reference interpreter. It includes work across the phases and explicit fallback
for the parts not implemented. Closures, callback collection methods, classes,
module execution and scope-sensitive host expressions still use the whole AST
program. Their benchmark results must not be presented as bytecode acceleration.

## Try it

Import all `src/*.cls`, including the two new dependencies `ASF_Bytecode` and
`ASF_FastVM`. This branch also includes the earlier native thunk classes.

```vb
Dim engine As New ASF, program As Long
engine.FastExecution = True
program = engine.Compile("let s=0;for(let i=0;i<1000;i+=1){s+=i;};return s;")
Debug.Print engine.Run(program)
Debug.Print engine.LastExecutionBackend, engine.FastFallbackReason
```

`FastExecution` defaults to False. Enabling it disables `NativeExecution` and
vice versa. Disabling either leaves the AST mode unless the other is enabled.
`LastExecutionBackend` is `Array bytecode`, `VBA fallback`, `VBA AST`, or `C thunk`.
`FastRunCount`/`FastFallbackCount` count invocations, not opcodes or test cases.
`FastOptimizations = False` disables constant folding, fused instructions and
map-position caching together; the slot/array/frame architecture remains.
Changing it clears the compiled-program cache. `Compile` also clears that cache.
Call `ClearFastCache` after external mutation through `GetGlobals`.

## Architecture and phase coverage

1. **Expressions, variables, control flow and named functions:** lowered to
   parallel `Long()` opcode/operand/handler arrays and a Variant constant pool.
   Compile-time interned names index numeric runtime bindings. ASF assigns to
   the nearest *existing caller* binding; these are therefore dynamic slot
   bindings, not JavaScript lexical depth/slot resolution. A single dispatch
   loop uses a preallocated, growing Variant stack and explicit return frames.
   Named calls do not recurse through VBA. A separate 1,000-deep recursion
   check verifies frame growth; the AST engine exhausts the VBA stack much
   earlier, so deep recursion is an intentional capability difference.
2. **Closures, arrays and objects:** literals, reads, object spread, array
   spread, nested indexing and writes are supported. Map reads cache a key's
   position at each instruction, checking that key on every hit so different
   shapes and removals cannot produce stale reads. Writes use ASF's existing
   lvalue services with already evaluated operands. **Closures/shared cells and
   callback invocation are not lowered.** Clearing an arena at end-of-run would
   invalidate escaped closures, and replacing `ShareFrames` with copied values
   would break shared writes. This branch preserves these through preflight
   fallback instead of introducing an incompatible closure representation.
3. **Builtins, templates, switch, iteration and spread/rest/destructuring:**
   numeric IDs select common string operations and the supported builtins
   `range`, `flatten`, `clone`, `isArray`, `isNumeric`. String primitives and
   array length execute directly; other builtin services remain host calls.
   `switch`, `for-in/of`, template strings, named-function rest parameters and
   array destructuring lower to bytecode. Array spread and clone/flatten
   preserve the reference engine's repeated evaluation. Call-argument spread,
   Math and most collection methods remain fallback. There is no reusable Math
   template: eligible scalar runs simply avoid creating Math/scope Maps.
4. **Classes, this/super and try/catch:** per-PC handler targets implement
   try/catch, including nested handlers and reference function error unwinding.
   VBA arithmetic/coercion errors still enter one dispatch-level error handler;
   this is not an exception-free numeric kernel. **Classes and this/super remain
   fallback.** Their constructor/method environments, bound receivers and
   inherited dispatch cannot safely reuse the current named-call frame model.
   Control transfers inside try/switch that are not yet lowered also reject.
5. **Modules, prototypes, collection overrides and host expressions:** ordinary
   Collection property/method services work and respect AppAccess and
   OverrideCollMethods. Explicit modules, `prototype.COM`, registered prototype
   state, arbitrary host function calls and `@(...)` remain fallback. Host
   expressions can see the whole ASF scope, which numeric slot frames do not
   currently expose. A literal-operand bridge must not pretend to provide that
   scope. Injected named bindings and opaque object arguments also fall back.

Other reasons for fallback include verbose/call-trace mode, nonempty raw scope
initializers, reserved runtime names, named functions used as values or shadowed,
and unsupported AST nodes. The reachable named-function graph is checked before
execution, including branches not taken at runtime. Unsupported programs never
start the bytecode interpreter; a runtime error never causes AST replay.
Argument coercion occurs only after eligibility is established.

The new backend is VBA, independent of the x64 C thunk. It calls existing ASF
services at explicit boundaries; this is not a claim that every bytecode opcode
is free of Maps/Collections. The hot scalar loop and named-call dispatch are.
Temporary operand nodes for lvalue/COM services hold evaluated values, not the
program's unevaluated expression tree. Cached programs hold no per-invocation
stack or scope. The default AST interpreter is unchanged apart from backend
selection/diagnostics; the map cache helper is used only by the new backend.

## Validation and performance

Run from the repository root with Python, pywin32, psutil and desktop Excel:

```powershell
python Tools/thunk/run_excel.py --output Tools/thunk/results/fast-tests
python Tools/thunk/run_excel.py --garcia --output Tools/thunk/results/garcia
python Tools/bytecode/report.py Tools/thunk/results/garcia Tools/bytecode/results/2026-11-03
python libreoffice/tools/build.py
```

The supervisor follows stdVBA/AGENTS.md: owned hidden Excel, disposable workbook,
CRLF verification before every import, owned-dialog watchdog, bounded execution,
and exact AccessVBOM/error-trapping registry restoration. The source default
backend is changed only in temporary test copies. Route accounting distinguishes
real lowered runs from fallback and from calls with no valid program.

The existing 255-test suite runs in all three modes. One known nested-array
VBAExpressions failure (`host_sees_array_mutation`, 42 expected / 8 actual)
reproduces on the original source plus the necessary private ByVal compile repair.
It is not fixed or suppressed here. Targeted tests additionally check coercions,
error/output preservation, evaluation order, shared object mutation, array copy
semantics, rest/spread/destructuring, cache changes, backend switching, collection
overrides, fallback side effects and explicit frame growth.

The benchmark includes the original eight thunk workloads, all eleven upstream
`AsfBench` workloads (including three intentional overlaps), and three synthetic
materials/payroll/error workloads. It contains no proprietary SGBO code or data.
QPC measures calibrated VBA batches, seven rounds with rotating engine order,
three phases: creation+Compile, creation+Compile+first Run, and warmed Run.
The bytecode lowering cost is in first Run, not Compile. Every batch validates
its answer before and after timing, and records the actual backend/fallback
reason. Ordinary Variant locals materialize ParamArray arguments in all engines,
avoiding a pre-existing ASF array-element wrapper problem on this Office host.
The supplementary ablation compares the combined optimization toggle, not the
individual causal effect of each optimization. Timings are descriptive medians
on one machine, not a statistical claim or a cross-host ranking.

LibreOffice is not installed on this machine. Its existing AST distribution
still builds; the generator deliberately excludes both experimental backends.
No LibreOffice runtime benchmark or 32-bit Office validation is claimed.

See `results/2026-11-03/report.html`, `samples.csv`, `summary.json` and
`results.json` for measured times, routing, raw samples and source hashes.
