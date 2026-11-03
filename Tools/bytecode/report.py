"""Summarize the three-engine Excel measurements; no third-party packages."""
from pathlib import Path
import argparse,csv,json,statistics,shutil

def write(path,text):
    path.write_bytes(text.replace('\r\n','\n').replace('\n','\r\n').encode('utf8'))

def main(source,output):
    output.mkdir(parents=True,exist_ok=True)
    result=json.loads((source/'results.json').read_text())
    assert result['status']=='passed'
    bench=result['benchmark']; rows=[]
    for case,name in enumerate(bench['cases']):
        for phase,label in enumerate(bench['phases']):
            row=dict(case=case,name=name,phase=phase,phase_name=label,engines=[])
            for engine,engine_name in enumerate(bench['engines']):
                values=[s['seconds']/s['count']*1e6 for s in bench['samples'] if (s['case'],s['phase'],s['engine'])==(case,phase,engine)]
                assert len(values)==7
                row['engines'].append(dict(name=engine_name,median_us=statistics.median(values),min_us=min(values),max_us=max(values),**bench['routes'][f'{engine}:{case}']))
            for e in row['engines']: e['speedup']=row['engines'][0]['median_us']/e['median_us']
            rows.append(row)
    ablation=[]
    for case in sorted({s['case'] for s in bench['ablation']}):
        medians=[statistics.median(s['seconds']/s['count']*1e6 for s in bench['ablation'] if s['case']==case and s['engine']==engine) for engine in (2,3)]
        ablation.append(dict(name=bench['cases'][case],optimized_us=medians[0],unoptimized_us=medians[1],speedup=medians[1]/medians[0]))
    write(output/'summary.json',json.dumps(dict(rows=rows,ablation=ablation),indent=2))
    write(output/'results.json',json.dumps(result,indent=2))
    with (output/'samples.csv').open('w',newline='',encoding='utf8') as f:
        w=csv.writer(f);w.writerow(['round','workload','phase','engine','count','seconds','us_per_call','route','fallback_reason'])
        for s in bench['samples']:
            route=bench['routes'][f"{s['engine']}:{s['case']}"]
            w.writerow([s['round']+1,bench['cases'][s['case']],bench['phases'][s['phase']],bench['engines'][s['engine']],s['count'],s['seconds'],s['seconds']/s['count']*1e6,route['route'],route['reason']])
    for filename in ('settings-before.json','settings-restored.json','dialogs.json','fast-routes.log','fast.log','native.log','vba.log'):
        if (source/filename).exists(): shutil.copy2(source/filename,output/filename)
    page=HTML.replace('ROWS_DATA',json.dumps(rows)).replace('ABLATION_DATA',json.dumps(ablation)).replace('PROVENANCE_DATA',json.dumps(dict(date=result['measured_at_utc'],office=result['office'],commit=result['commit'],tests=result['fast_differential'],routes=result['fast_routes'])))
    write(output/'report.html',page)
    for r in rows:
        if r['phase']==2:
            print(r['name']+': '+', '.join(f"{e['name']} {e['median_us']:.2f} us ({e['speedup']:.2f}x; {e['route']})" for e in r['engines']))

HTML='''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>ASF three-engine benchmark</title>
<style>
:root{color-scheme:light dark;font-family:var(--font-sans,system-ui);color:var(--foreground,CanvasText)}
body{margin:0;font-size:14px;line-height:1.5}p{margin:8px 0}label{display:inline-block;margin:10px 15px 10px 0}select{font:inherit;background:var(--secondary,Canvas);color:inherit;border:1px solid var(--border,GrayText);border-radius:var(--radius,4px);padding:4px}svg{width:100%;display:block}svg text{fill:var(--foreground,CanvasText);font:11px var(--font-sans,system-ui)}.muted{font-size:12px;color:var(--muted-foreground,GrayText)}.native{color:var(--chart-1,SteelBlue)}.fast{color:var(--chart-2,SeaGreen)}.legend{display:flex;gap:18px;flex-wrap:wrap;font-size:12px}.legend span:before{content:' ';display:inline-block;width:12px;height:8px;margin-right:5px;background:currentColor}table{border-collapse:collapse;width:100%;font-size:12px}td,th{padding:5px;text-align:right;border-bottom:1px solid var(--border,GrayText)}td:first-child,th:first-child{text-align:left}details{margin:12px 0}.scroll{overflow-x:auto}code{font-size:12px}
</style></head><body>
<p><strong>Array bytecode accelerates lowered workloads; the five-phase rewrite remains incomplete.</strong> Closures, callback pipelines and classes below still execute the VBA AST program. Hollow bars mark fallback.</p>
<label>Phase <select id="phase"><option value="2">Warm Run</option><option value="1">Create + Compile + first Run</option><option value="0">Create + Compile</option></select></label>
<label>Workloads <select id="group"><option value="upstream">Upstream AsfBench (11)</option><option value="original">Earlier thunk benchmark (8)</option><option value="synthetic">Synthetic workloads (3)</option><option value="all">All workloads (22)</option></select></label>
<div class="legend"><span class="native">C thunk</span><span class="fast">Array bytecode / fallback</span><small>VBA AST = 1x baseline</small></div>
<svg id="chart" role="img" aria-label="Median execution speedup relative to VBA AST; logarithmic horizontal scale"></svg>
<p class="muted">Horizontal axis: median VBA AST time / median engine time (logarithmic scale; higher is faster). Vertical axis: workload. Seven rotating rounds, QPC batches calibrated to 30 ms, capped at 16,384 iterations. Descriptive results from one host; not a significance test.</p>
<p class="muted" id="source"></p>
<details><summary>Median times, actual routing and fallback reasons</summary><div class="scroll"><table><thead><tr><th>Workload</th><th>AST (us)</th><th>Thunk (us)</th><th>Bytecode mode (us)</th><th>Bytecode route / reason</th></tr></thead><tbody id="values"></tbody></table></div></details>
<details><summary>Combined optimization ablation (warm runs)</summary><p class="muted">Folding, fused instructions and map-position caching on versus off, both using slot bindings and bytecode. Seven rotating rounds. These measurements cannot assign causality to an individual optimization.</p><div class="scroll"><table><thead><tr><th>Workload</th><th>On (us)</th><th>Off (us)</th><th>Off / on</th></tr></thead><tbody id="ablation"></tbody></table></div></details>
<details><summary>Validation and limitations</summary><p id="tests"></p><p>All three modes pass 254/255 upstream tests. The same pre-existing nested-array VBAExpressions failure remains (42 expected / 8 actual). Routing counts are invocations, not test cases. Many upstream tests request verbose tracing and therefore use fallback. Separate targeted tests assert actual lowering as well as value/type parity.</p><p>Compile measures AST parsing. Bytecode lowering occurs on first Run and is included in the cold execution phase. Validation and Python/COM communication are outside timed batches. Input arguments are materialized identically in all modes. Synthetic materials/payroll workloads contain no proprietary SGBO data. LibreOffice's AST distribution builds; no LibreOffice runtime or 32-bit Office result is claimed. Exact sources, raw samples and Office-setting restoration records accompany this report.</p></details>
<script>
const rows=ROWS_DATA,ablation=ABLATION_DATA,provenance=PROVENANCE_DATA,ns='http://www.w3.org/2000/svg';
const cell=(tr,value)=>{const td=document.createElement('td');td.textContent=value;tr.append(td);};
function draw(){const phase=+document.getElementById('phase').value,group=document.getElementById('group').value;
const data=rows.filter(r=>r.phase===phase&&(group==='all'||group==='original'&&r.case<8||group==='upstream'&&r.case>=8&&r.case<19||group==='synthetic'&&r.case>=19));
const svg=document.getElementById('chart'),h=70+data.length*42;svg.replaceChildren();svg.setAttribute('viewBox',`0 0 728 ${h}`);svg.style.height=h+'px';
const add=(tag,attrs,text)=>{const e=document.createElementNS(ns,tag);for(const[k,v]of Object.entries(attrs))e.setAttribute(k,v);if(text!==undefined)e.textContent=text;svg.append(e);return e;};
const min=Math.min(-2,Math.floor(Math.log2(Math.min(...data.flatMap(r=>r.engines.slice(1).map(e=>e.speedup)))))),max=Math.max(1,Math.ceil(Math.log2(Math.max(...data.flatMap(r=>r.engines.slice(1).map(e=>e.speedup))))));
const x=v=>175+(Math.log2(v)-min)/(max-min)*475,base=x(1);
for(let n=min;n<=max;n++){const xx=x(2**n);add('line',{x1:xx,x2:xx,y1:10,y2:h-43,stroke:'var(--border,GrayText)','stroke-dasharray':n===0?'0':'2 4'});add('text',{x:xx,y:h-25,'text-anchor':'middle'},(2**n)+'x');}
data.forEach((r,i)=>{const y=17+i*42;add('text',{x:167,y:y+15,'text-anchor':'end'},r.name);r.engines.slice(1).forEach((e,j)=>{const xx=x(e.speedup),color=j===0?'var(--chart-1,SteelBlue)':'var(--chart-2,SeaGreen)',fallback=e.route==='VBA fallback',yy=y+j*15;const bar=add('rect',{x:Math.min(xx,base),y:yy,width:Math.max(2,Math.abs(xx-base)),height:10,fill:fallback?'none':color,stroke:color,'stroke-dasharray':fallback?'3 2':'0'});const title=document.createElementNS(ns,'title');title.textContent=`${e.name}: ${e.median_us.toFixed(2)} us/call; ${e.route}; ${e.reason}`;bar.append(title);add('text',{x:xx+(e.speedup>=1?4:-4),y:yy+9,'text-anchor':e.speedup>=1?'start':'end'},e.speedup.toFixed(2)+'x'+(fallback?' F':''));});});
add('text',{x:412,y:h-5,'text-anchor':'middle'},'Speedup versus VBA AST (ratio; higher is faster)');
const tbody=document.getElementById('values');tbody.replaceChildren();data.forEach(r=>{const tr=document.createElement('tr');[r.name,...r.engines.map(e=>e.median_us.toFixed(2)),r.engines[2].route+(r.engines[2].reason?': '+r.engines[2].reason:'')].forEach(v=>cell(tr,v));tbody.append(tr);});}
for(const r of ablation){const tr=document.createElement('tr');[r.name,r.optimized_us.toFixed(2),r.unoptimized_us.toFixed(2),r.speedup.toFixed(3)+'x'].forEach(v=>cell(tr,v));document.getElementById('ablation').append(tr);}
document.getElementById('source').textContent=`Source: local Excel ${provenance.office.version} x64, build ${provenance.office.build}; ${provenance.date.slice(0,10)} UTC. Source commit ${provenance.commit.slice(0,12)}; exact imported-source hashes in results.json.`;
document.getElementById('tests').textContent=`Targeted bytecode checks: ${provenance.tests}; native regression assertions: 723. Upstream invocations: ${provenance.routes} (lowered / fallback / no valid program).`;
document.getElementById('phase').addEventListener('change',draw);document.getElementById('group').addEventListener('change',draw);draw();
</script></body></html>'''
if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('source',type=Path);parser.add_argument('output',type=Path);args=parser.parse_args();main(args.source,args.output)
