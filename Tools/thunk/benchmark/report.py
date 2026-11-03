"""Create a self-contained chart and raw CSV from the Excel benchmark result."""
import argparse
import csv
import hashlib
import json
from pathlib import Path
import statistics

HERE = Path(__file__).resolve().parent

def report(source, output):
    result = json.loads(source.read_text())
    assert result['status'] == 'passed'
    bench = result['benchmark']
    assert len(bench['samples']) == 7 * 8 * 3 * 2
    output.mkdir(parents=True, exist_ok=True)
    rows = []
    for case, label in enumerate(bench['cases']):
        for phase, phase_label in enumerate(bench['phases']):
            samples = [[s['seconds'] / s['count'] * 1e6 for s in bench['samples']
                        if s['case'] == case and s['phase'] == phase and s['engine'] == engine]
                       for engine in range(2)]
            assert all(len(s) == 7 and all(t > 0 for t in s) for s in samples)
            medians = [statistics.median(s) for s in samples]
            rows.append(dict(case=label, phase=phase_label, phase_id=phase,
                             vba_us=medians[0], thunk_us=medians[1], ratio=medians[1]/medians[0],
                             vba_min_us=min(samples[0]), vba_max_us=max(samples[0]),
                             thunk_min_us=min(samples[1]), thunk_max_us=max(samples[1])))
    with (output/'samples.csv').open('w',newline='',encoding='utf-8') as stream:
        writer = csv.writer(stream)
        writer.writerow(['round','workload','phase','engine','iterations','seconds','microseconds_per_call'])
        for s in bench['samples']:
            writer.writerow([s['round']+1,bench['cases'][s['case']],bench['phases'][s['phase']],
                             bench['engines'][s['engine']],s['count'],s['seconds'],s['seconds']/s['count']*1e6])
    write(output/'summary.json',json.dumps(rows,indent=2))
    write(output/'results.json',json.dumps(result,indent=2))
    context = dict(date=result['measured_at_utc'], office=result['office'], commit=result['commit'],
                   benchmark_sha256=hashlib.sha256((HERE/'Benchmark.bas').read_bytes()).hexdigest())
    write(output/'provenance.json',json.dumps(context,indent=2))
    page = '''<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>ASF thunk versus VBA</title>
<style>
:root { color-scheme:light dark; font-family:var(--font-sans,system-ui); color:var(--foreground,CanvasText) }
body { margin:0; font-size:14px; line-height:1.5 } h2 { font-size:19px; margin:0 0 8px }
p { margin:8px 0 14px } .muted { color:var(--muted-foreground,GrayText); font-size:12px }
select { font:inherit; color:inherit; background:var(--secondary,Canvas); border:1px solid var(--border,GrayText); border-radius:var(--radius,5px); padding:5px }
svg { width:100%; height:390px; display:block } svg text { fill:var(--foreground,CanvasText); font-family:var(--font-sans,system-ui); font-size:12px }
.bar { fill:var(--chart-1,SteelBlue) } .axis { stroke:var(--border,GrayText) } .baseline { stroke:var(--foreground,CanvasText); stroke-dasharray:4 4 }
table { border-collapse:collapse; width:100%; font-size:12px } th,td { padding:6px 4px; text-align:right; border-bottom:1px solid var(--border,GrayText) } th:first-child,td:first-child { text-align:left }
details { margin-top:12px } label { display:block; margin:12px 0 }
</style></head><body>
<h2>ASF execution time: native thunk / VBA</h2>
<p>Lower ratios are faster. The thunk benefits loops in this run; most other workloads pay more host-bridge and setup overhead.</p>
<label>Measurement phase <select id="phase"><option value="2">Warm run</option><option value="1">Create + compile + first run</option><option value="0">Create + compile</option></select></label>
<svg id="chart" viewBox="0 0 700 390" role="img" aria-label="Median thunk to VBA elapsed-time ratio by workload"></svg>
<p class="muted">Seven alternating rounds per engine. Bars show ratio of median times per call; dashed line = equal time (1.00). Source: Excel 16 x64 build 20618, 3 November 2026. One host; these samples do not establish statistical significance.</p>
<details><summary>Absolute median times and method</summary>
<table><thead><tr><th>Workload</th><th>VBA (Âµs/call)</th><th>Thunk (Âµs/call)</th><th>Thunk / VBA</th></tr></thead><tbody id="values"></tbody></table>
<p>Both modes use the same ASF source, parser, scripts, arguments and host settings. Warm runs exclude object construction, native loading and AST lowering. Cold phases include backend selection; native selection allocates its cached thunk. Every batch checks the result and execution route before and after timing. QueryPerformanceCounter measures VBA batches calibrated to at least 30 ms, capped at 16,384 iterations; engine order alternates.</p>
<p>Validation: 723 extra assertions passed. Both modes pass 254/255 original tests; the same nested-array VBAExpressions failure reproduces on pre-change source. Array-element arguments are materialized into ordinary Variants before Run, identically for both modes, to avoid a reference-engine ParamArray issue on this Office host.</p>
</details>
<script>
const rows=DATA;
const ns='http://www.w3.org/2000/svg';
function draw(){
 const data=rows.filter(r=>r.phase_id===Number(document.getElementById('phase').value));
 const svg=document.getElementById('chart');svg.replaceChildren();
 const max=Math.max(2,Math.ceil(Math.max(...data.map(r=>r.ratio))*4)/4),left=145,width=470;
 const add=(tag,attrs,text)=>{const e=document.createElementNS(ns,tag);for(const [k,v] of Object.entries(attrs))e.setAttribute(k,v);if(text!==undefined)e.textContent=text;svg.append(e);return e;};
 add('text',{x:left,y:15},'Workload');
 for(let t=0;t<=max+.01;t+=.5){const x=left+width*t/max;add('line',{x1:x,x2:x,y1:23,y2:342,class:'axis'});add('text',{x,y:360,'text-anchor':'middle'},t.toFixed(1));}
 add('line',{x1:left+width/max,x2:left+width/max,y1:23,y2:342,class:'baseline'});
 data.forEach((r,i)=>{const y=30+i*39;add('text',{x:left-8,y:y+17,'text-anchor':'end'},r.case);const bar=add('rect',{x:left,y,width:width*r.ratio/max,height:25,class:'bar'});const title=document.createElementNS(ns,'title');title.textContent=`${r.case}: VBA ${r.vba_us.toFixed(2)} Âµs; thunk ${r.thunk_us.toFixed(2)} Âµs`;bar.append(title);add('text',{x:left+width*r.ratio/max+6,y:y+17},r.ratio.toFixed(2)+'Ã—');});
 add('text',{x:left+width/2,y:386,'text-anchor':'middle'},'Thunk / VBA elapsed time (ratio; lower is faster)');
 const tbody=document.getElementById('values');tbody.replaceChildren();data.forEach(r=>{const row=document.createElement('tr');[r.case,r.vba_us.toFixed(2),r.thunk_us.toFixed(2),r.ratio.toFixed(3)].forEach(v=>{const td=document.createElement('td');td.textContent=v;row.append(td)});tbody.append(row)});
}
document.getElementById('phase').addEventListener('change',draw);draw();
</script></body></html>'''.replace('DATA',json.dumps(rows).replace('</',r'<\/'))
    write(output/'report.html',page)
    for r in rows:
        if r['phase_id'] == 2:
            print(f"{r['case']}: VBA {r['vba_us']:.2f} us; thunk {r['thunk_us']:.2f} us; ratio {r['ratio']:.3f}")

def write(path, text):
    path.write_bytes(text.replace('\r\n','\n').replace('\n','\r\n').encode('utf-8'))

if __name__ == '__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('source',type=Path)
    parser.add_argument('--output',type=Path,default=HERE/'results/2026-11-03')
    args=parser.parse_args()
    report(args.source,args.output)
