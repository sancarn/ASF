# Garcia bytecode backend experiment

This branch implements feature coverage across the five phases proposed in
[WS Garcia's PR comment](https://github.com/sancarn/ASF/pull/1#issuecomment-6600185041).
It provides two execution paths within the opt-in VBA bytecode backend. Compact
slot bytecode handles eligible programs; shared-scope bytecode handles closures,
classes, modules and host integration. The earlier whole-program AST fallback
has been removed. Complex operations still call ASF's semantic services, which
execute their child script expressions and function bodies through bytecode.
This is an experimental implementation, not a claim that every operation has
become a primitive opcode or that the proposal's performance targets are met.

## Try it

Import all `src/*.cls`, including the three new dependencies `ASF_Bytecode`,
`ASF_FastVM` and `ASF_ScopeVM`. This branch includes the earlier native thunk.

```vb
Dim engine As New ASF, program As Long
engine.FastExecution = True
program = engine.Compile("let f=fun(n){return n+1;};return f(4);")
Debug.Print engine.Run(program)
Debug.Print engine.LastExecutionBackend, engine.FastScopeEntryCount
```

`FastExecution` defaults to False and is mutually exclusive with
`NativeExecution`. The routes are `Array bytecode`, `Scope bytecode`, `VBA AST`
and `C thunk`. `FastFallbackCount` and `FastFallbackReason` are retained for
existing experimental callers; successful bytecode runs have no fallback.
An execution error never retries the program through the AST interpreter.

`FastOptimizations = False` disables compact constant folding, fused instructions
and map-position caching together. It leaves both bytecode architectures in
place; it does not toggle scope execution or its cell cache. Compile, module
compilation, changing Globals and changing the optimization flag invalidate
compiled caches. Call `ClearFastCache` after external AST/function-table changes
through `GetGlobals`.

`FastRunCount` counts completed invocations, including handled program errors.
`FastScopeEntryCount` counts shared-scope body/expression entries and direct call
frames. `FastHostServiceCount` counts semantic boundary calls, not instructions.
`FastClosureReleaseCount` counts isolated closure/bound-method cycles broken at
run completion. `FastScopeProgramCount` reports the current immutable code cache.
Counters are diagnostics, not numbers of test cases.

## Architecture and phase coverage

1. **Slots, bytecode, control flow and named functions.** Parallel Long opcode,
   operand and exception-handler arrays plus a Variant constant pool replace
   repeated statement dispatch. The compact executor uses numeric bindings,
   a growing operand stack and explicit return frames. ASF assigns into the
   nearest existing caller binding, so these are dynamic bindings rather than
   JavaScript lexical slots. Eligible runs avoid constructing scope/Math maps.
2. **Closures, arrays and objects.** Shared-scope variable sites resolve an owner
   frame and a numeric position in its Variant storage. Captures share that
   storage, preserving future declarations and writes through `ShareFrames`.
   Direct closure calls use the trampoline. Arrays retain ASF's copy semantics;
   objects retain reference semantics. Compact map reads use guarded position
   caches. The scope cache is invalidated across calls/services that can mutate
   bindings. Callback services execute compiled callback bodies. The arena
   breaks closure/environment and bound-method/instance cycles only when a run
   is isolated and returns a scalar. Object/array arguments, injection, imports,
   host access and escaping results conservatively preserve environments. This
   is not a general cycle collector; escaped environments are never cleared.
3. **Builtins, iteration and structured values.** Compact instructions cover
   templates, switch, for-in/of, spread/rest and destructuring. Shared-scope
   variants use existing semantic services with bytecode children. Twenty-one
   array-helper methods have numeric call-site IDs, while generic polymorphic
   and COM dispatch retain ASF's method services. Argument order and ASF's
   repeated evaluation in spread/clone/flatten are preserved. Scope runs retain
   fresh per-run Math state; the compact path does not allocate it.
4. **Classes, this/super and exceptions.** Constructors, inheritance, static and
   bound methods use existing class descriptors/receiver rules with compiled
   bodies. Direct bound calls on variable receivers use explicit frames when
   tracing is off. Constructors, super, complex receivers and tracing-sensitive
   dispatch remain semantic services; recursive chains through those services
   still consume VBA stack. Both bytecode paths use per-PC catch targets. They
   preserve ASF's ignored catch binding, consumed break/continue and silent
   classic-for condition failures. Closure failures retain the reference
   environment frame, including its observable effect on recursive closures.
5. **Modules, prototypes and host integration.** Module imports/exports and
   caches, `prototype.COM`, collection overrides, injected variables and
   `@(...)` run with the real shared scope. Module and host services enter
   compiled script bodies. Prototype precedence and AppAccess checks remain in
   ASF's services. Tests exercise exported closures over repeated runs,
   persistent COM prototypes, host-visible mutation and same-engine reentry.

`ASF_Bytecode` emits compact or scoped instructions; `ASF_FastVM` owns immutable
code and executes the compact path; `ASF_ScopeVM` executes scoped code and its
return frames. Code caches retain their AST roots so pointer-keyed entries
cannot alias recycled nodes. VM-created operand descriptors are transient and
are never cached with their runtime values. A repeated class/callback test
checks cache size remains stable. Invocation stacks and cell owners are local.

The scope path is deliberately compatible with ASF's existing frame maps and
semantic services. Its allocation and dispatch costs can outweigh bytecode
savings. The backend is VBA and independent of the x64 C thunk: invoking the
shared Variant arithmetic helper does not enter native C. The reference AST
remains the default. These tradeoffs are visible in the three-engine benchmarks.

## Validation and performance

From the repository root with Python, pywin32, psutil and desktop Excel:

```powershell
python Tools/thunk/run_excel.py --output Tools/thunk/results/fast-tests
python Tools/thunk/run_excel.py --garcia --output Tools/thunk/results/garcia-complete
python Tools/bytecode/report.py Tools/thunk/results/garcia-complete Tools/bytecode/results/2026-11-03-scope
python libreoffice/tools/build.py
```

The supervisor follows stdVBA/AGENTS.md: owned hidden Excel, disposable workbook,
CRLF verification before imports, owned-dialog watchdog, bounded execution and
exact AccessVBOM/error-trapping restoration. Benchmark runs have a 480-second
limit; ordinary validation has 240 seconds. Test defaults change only in the
imported temporary copies. Scope-route checks require actual bytecode entries
and zero whole-program fallbacks.

Validation covers 192 targeted bytecode assertions, 473 Variant/log/reentry
checks against the AST and 725 native regression assertions. Each of the three
engines passes 254 of 255 upstream tests. The known nested-array host-expression
failure (`host_sees_array_mutation`, expected 42 / actual 8) reproduces on the
original source with only its necessary private ByVal compile repair. It is not
suppressed. Upstream bytecode invocation routes are 50 compact / 209 scoped /
0 fallback / 1 invalid-program call. Additional checks include escaped closures
across runs, finalizer timing and default-property call counts, exception-frame
and parameter-binding error behavior, cache stability, 1,000-deep direct closure
calls and 500-deep methods.

The benchmark includes eight earlier thunk workloads, all eleven upstream
`AsfBench` workloads (three intentionally overlap), and three synthetic
materials/payroll/error workloads. No proprietary SGBO data is used. QPC times
calibrated batches over seven rounds with rotating engine order. Phases separate
creation+Compile, creation+Compile+first Run, and warmed Run. Lowering occurs on
first Run. Every batch validates answers and routes outside the timer. Ordinary
Variant locals materialize ParamArray arguments identically in all engines.
The ablation toggles folding/fusion/map caching together; it cannot identify
individual optimization effects. Results are descriptive measurements on one
host, not a cross-host or statistical performance guarantee.

On this host, warm compact bytecode improves the upstream arithmetic loop
16.71x, recursive calls 28.17x and template literals 20.00x versus AST. Scope
bytecode does not improve these measured feature workloads: closure speedup is
0.91x, classes 0.86x, array methods 0.75x and the small payroll pipeline 0.46x
(about 2.18x as long as AST). Feature coverage therefore does not imply a
performance win. The report separates these routes and cold startup costs.

The new results are in `results/2026-11-03-scope/`: interactive `report.html`,
`samples.csv`, `summary.json`, `results.json`, logs and setting-restoration records.
The earlier `results/2026-11-03/` archive measures the partial implementation and
is retained for provenance; its fallback results do not describe this revision.

LibreOffice is not installed here. Its existing ten-module AST distribution
builds with experimental classes and helpers excluded. No LibreOffice runtime
or 32-bit Office benchmark is claimed.
