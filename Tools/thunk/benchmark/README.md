# ASF thunk versus VBA: Excel comparison

This branch adds a repeatable comparison to `feature/thunk-vm`. It does not alter
the VM implementation. Both modes use runtime commit `13fa404`, the same source
scripts and arguments, and the same Excel process. `NativeExecution=False` selects
the reference VBA AST interpreter; `True` selects the experimental C backend.

Open [the interactive report](results/2026-11-03/report.html) to switch between
warm execution and the two cold phases. [Raw samples](results/2026-11-03/samples.csv),
[full run metadata](results/2026-11-03/results.json), [medians and ranges](results/2026-11-03/summary.json)
and [source provenance](results/2026-11-03/provenance.json) are checked in.

## Result

On 3 November 2026, Excel 16 x64 build 20618 on an Intel Core i7-9700KF:

- The branch-heavy while loop takes 31.37 ms per warm native call versus 46.35 ms
  in VBA: **32.3% less elapsed time** (1.48x throughput).
- The arithmetic for loop takes 74.39 ms versus 81.29 ms: **8.5% less time**.
- The other six warm workloads are slower in native mode: about 5% for the array
  callback, 6% for the predicate, 10% for arithmetic, 21% for Fibonacci, 33% for
  COM collections and 64% for strings.
- Creating/configuring the native engine and compiling its AST costs more in all
  workloads. The backend caches code per engine, so this cost is amortized only
  when callers retain and reuse their engine.

This supports keeping native execution opt-in. ASF's AST traversal can benefit
from native control flow, but scope lookups, recursive calls and complex values
still cross VBA bridges. It does not show a general performance improvement.
Seven rounds on one desktop host are limited evidence, not a significance test
or a prediction for other Office versions or workloads.

## Reproduce

Use Windows x64 Excel and the Python dependencies described in
[the implementation guide](../README.md). From the repository root:

```powershell
python Tools/thunk/run_excel.py --benchmark --output Tools/thunk/results/benchmark
python Tools/thunk/benchmark/report.py Tools/thunk/results/benchmark/results.json --output Tools/thunk/benchmark/results/new-run
```

Run only one Office supervisor at a time. The harness imports normalized CRLF
copies into a disposable hidden Excel process, watches only that process for
modal dialogs, enforces a deadline and restores the exact original AccessVBOM
and VBE error-trapping settings. Restoration evidence is included with the run.

The run first passes 723 extra assertions and compares the existing 255-test
suite under both engines. Both pass 254 tests; the same pre-existing
`host_sees_array_mutation` failure is recorded in both. No benchmark result was
accepted without checking its expected answer and native/VBA routing.

## Measurement details

`Benchmark.bas` defines eight workloads and three phases:

1. **Create + compile**: construct an engine, select its backend, enable the same
   host access and compile source to an AST. Selecting native mode loads its code;
   AST lowering remains lazy and is not included in this phase.
2. **Create + compile + first run**: the same setup plus one execution. This
   includes native lowering and other first-use initialization.
3. **Warm run**: reuse an engine after one validated warm-up. Construction,
   compilation, code loading and lowering are outside the timed interval.

QueryPerformanceCounter measures batches in VBA. Counts are calibrated separately
per workload, phase and engine to at least 30 ms, with a 16,384-call cap. Seven
rounds alternate engine order; the report divides batch time by count and takes
the median, then computes native median / VBA median. There is no result cache.
The last answer and execution route are checked after every measured batch;
validation and Python/Excel COM round-trip costs are outside the timed interval.
Small, identical VBA call-loop overhead remains inside both modes.

Arguments are materialized into ordinary Variant variables before timing. Passing
`args(0)` directly into the existing ASF `Run` ParamArray returned pointer-like
values in the **VBA** engine on this Office host, whereas an ordinary local holding
the same value returned the correct answer. The benchmark uses the working
calling convention identically for both engines. This is a harness workaround;
it does not change or claim to fix that reference-engine behavior.

The pure mode is the reference interpreter on the implementation branch, including
its backend-selection guard and private ByVal compilation repair. It is not a
separately renamed older release. Raw runtime hashes identify the exact code.
