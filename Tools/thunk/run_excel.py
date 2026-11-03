"""Isolated Excel tests, following stdVBA/AGENTS.md.

Requires desktop Excel, Python, pywin32 and psutil. The parent owns registry
restoration, modal detection and deadlines even if the COM worker hangs/crashes.
See stdVBA-LICENSE.txt for the supervisor cleanup adapted from stdVBA's harness.
"""
from pathlib import Path
import argparse
import ctypes
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import winreg

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]

def write(path, text):
    path.write_bytes(text.replace('\r\n', '\n').replace('\n', '\r\n').encode('utf-8'))

def worker(folder, smoke, baseline=False, benchmark=False):
    import win32com.client
    import win32process
    app = book = None
    result = {}
    try:
        candidate = win32com.client.DispatchEx('Excel.Application')
        pid = win32process.GetWindowThreadProcessId(candidate.Hwnd)[1]
        if pid in json.loads((folder/'original-pids.json').read_text()):
            raise RuntimeError('Excel reused an existing process; refusing to modify it')
        app = candidate
        (folder/'pid').write_text(str(pid))
        app.Visible = False
        app.DisplayAlerts = False
        book = app.Workbooks.Add()
        while book.Worksheets.Count < 3:
            book.Worksheets.Add()
        book.SaveAs(str(folder/'test.xlsm'), 52)
        result['office'] = dict(version=app.Version, build=app.Build, os=app.OperatingSystem)
        project = book.VBProject
        for path in sorted(folder.glob('*.cls')) + sorted(folder.glob('*.bas')):
            data = path.read_bytes().replace(b'\r\n', b'')
            assert b'\r' not in data and b'\n' not in data, str(path)
            project.VBComponents.Import(str(path))
        if baseline:
            result['vba'] = app.Run('ThunkProbe.ExistingSuite', False)
            result['status'] = 'baseline'
            print('Pre-change source: '+result['vba'], flush=True)
            return
        result['smoke'] = app.Run('ThunkProbe.Smoke')
        print(result['smoke'], flush=True)
        if not result['smoke'].startswith('35:'):
            raise RuntimeError('Native smoke test: '+result['smoke'])
        if not smoke:
            regression = app.Run('ThunkProbe.Regression')
            print('Regression: '+str(regression), flush=True)
            if str(regression).startswith('ERROR'): raise RuntimeError(regression)
            result['regression_checks'] = int(regression)
            for native in (False, True):
                label = 'native' if native else 'vba'
                print('Running '+label+' suite', flush=True)
                result[label] = app.Run('ThunkProbe.ExistingSuite', native)
                shutil.copy2(folder/'TestRunner.log', folder/(label+'.log'))
                print(label+': '+result[label], flush=True)
            failures = {label: [re.sub(r' \(\d+ ms\)$','',line) for line in (folder/(label+'.log')).read_text().splitlines() if line.startswith('FAIL ')] for label in ('vba','native')}
            known = ['FAIL host_VBA_Expressions/host_sees_array_mutation: Expected String <42>, actual String <8>']
            if result['vba'] != result['native'] or failures['vba'] != failures['native'] or failures['vba'] not in ([], known):
                raise RuntimeError('Existing suite parity/regression failed: '+str(failures))
            result['known_baseline_failures'] = failures['vba']
        if benchmark:
            cases = ['Arithmetic', 'Predicate', 'Strings', 'COM collection', 'For loop', 'While + branches', 'Fibonacci', 'Array callback']
            phases = ['Create + compile', 'Create + compile + first run', 'Warm run']
            result['benchmark'] = dict(cases=cases, phases=phases, samples=[], timer='QueryPerformanceCounter', rounds=7, target_batch_seconds=.03,
                                       engines=['VBA', 'Thunk'], cpu=__import__('platform').processor(), benchmark_source_sha256=hashlib.sha256((folder/'Benchmark.bas').read_bytes()).hexdigest())
            counts = {}
            def measure(engine, case, phase, count):
                value = str(app.Run('ThunkBenchmark.Measure', bool(engine), case, phase, count))
                if value.startswith('ERROR'): raise RuntimeError(f'engine={engine} case={case} phase={phase} count={count}: {value}')
                return float(value)
            for case in range(len(cases)):
                for phase in range(3):
                    for engine in range(2):
                        count = 1
                        while measure(engine,case,phase,count) < .03 and count < 16384:
                            count *= 2
                        counts[engine,case,phase] = count
            for round_ in range(7):
                for case in range(len(cases)):
                    for phase in range(3):
                        for offset in range(2):
                            engine = (round_ + case + phase + offset) % 2
                            count = counts[engine,case,phase]
                            seconds = measure(engine,case,phase,count)
                            result['benchmark']['samples'].append(dict(round=round_,case=case,phase=phase,engine=engine,count=count,seconds=seconds))
                print('Benchmark round '+str(round_+1)+'/7 complete',flush=True)
        result['status'] = 'passed'
    except Exception as exc:
        result.update(status='failed', error=repr(exc))
        raise
    finally:
        write(folder/'result.json', json.dumps(result, indent=2))
        try:
            if book is not None: book.Close(False)
        finally:
            if app is not None: app.Quit()

def prepare(folder, baseline=False, benchmark=False):
    for path in (ROOT/'src').glob('*.cls'):
        if baseline:
            if path.stem in ('ASF_Thunk', 'ASF_ThunkProgram'): continue
            source = subprocess.check_output(['git','show','527f172:src/'+path.name],cwd=ROOT).decode('cp1252').replace('\r\n','\n')
            # Necessary Excel compilation repair only, with no execution change.
            source = source.replace('                               stmts As Collection, evaluatedArgs As Collection, ByVal hasThis As Boolean',
                                    '                               ByVal stmts As Collection, evaluatedArgs As Collection, ByVal hasThis As Boolean')
        else:
            source = path.read_bytes().decode('cp1252').replace('\r\n', '\n')
        if path.stem == 'ASF' and not baseline:
            source = source.replace('    Set VM_.Globals = GLOBALS_\nEnd Sub',
                                    '    Set VM_.Globals = GLOBALS_\n    VM_.NativeExecution = ThunkProbe.NativeDefault\nEnd Sub', 1)
        # Preserve source encoding; the VBE imports ANSI modules.
        (folder/path.name).write_bytes(source.replace('\n','\r\n').encode('cp1252'))
    if benchmark: write(folder/'Benchmark.bas',(HERE/'benchmark/Benchmark.bas').read_text())
    runner = (ROOT/'test/TestRunner.bas').read_text(encoding='cp1252')
    runner += '\nPublic Function ThunkSuiteResult() As String\n ThunkSuiteResult = pPassed & "/" & pRun & "/" & pFailures.Count\nEnd Function\n'
    (folder/'TestRunner.bas').write_bytes(runner.replace('\n','\r\n').encode('cp1252'))
    probe = (HERE/'Probe.bas').read_text()
    if baseline: probe = probe[:probe.index('Public Function Smoke()')]
    write(folder/'Probe.bas', probe)
    if not baseline:
        for helper in HERE.glob('*.cls'): write(folder/helper.name, helper.read_text())
    # Test fixtures expected by the repository's existing suite.
    for path in (ROOT/'test').rglob('*.vas'):
        target = folder/path.relative_to(ROOT/'test')
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path,target)
    write(folder/'UserDefFunctions.bas', '''Attribute VB_Name = "UserDefFunctions"
Option Explicit
Public Function ThisWBname() As String
    ThisWBname = ThisWorkbook.Name
End Function
Public Function HostCheckSecondNestedElement(arr As Variant) As String
    HostCheckSecondNestedElement = CStr(arr(LBound(arr))(LBound(arr(LBound(arr))) + 1))
End Function
''')

def supervise(output, smoke=False, baseline=False, benchmark=False):
    import psutil
    import win32con
    import win32gui
    import win32process
    output.mkdir(parents=True, exist_ok=True)
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'Excel.Application\CurVer') as key:
        version = winreg.QueryValueEx(key, '')[0].rsplit('.',1)[1]+'.0'
    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        try:
            with winreg.OpenKey(hive, rf'Software\Policies\Microsoft\Office\{version}\Excel\Security') as key:
                if not winreg.QueryValueEx(key, 'AccessVBOM')[0]:
                    raise RuntimeError('Managed policy disables VBA project access')
        except FileNotFoundError: pass
    # VBA7.1 is the installed VBE used by Office 16. Refuse other versions until
    # their debugger key is established rather than changing a guessed setting.
    if version != '16.0': raise RuntimeError('Determine this Office version\'s VBE registry key first')
    settings = [(rf'Software\Microsoft\Office\{version}\Excel\Security','AccessVBOM',1),
                (r'Software\Microsoft\VBA\7.1\Common','BreakOnAllErrors',0),
                (r'Software\Microsoft\VBA\7.1\Common','BreakOnServerErrors',0)]
    saved, dialogs = [], []
    process = office_pid = None
    original_pids = {p.pid for p in psutil.process_iter()}
    started = time.time()
    with tempfile.TemporaryDirectory(prefix='asf-thunk-test-') as temp:
        folder = Path(temp)
        prepare(folder, baseline, benchmark)
        write(folder/'original-pids.json',json.dumps(sorted(original_pids)))
        try:
            for path, name, value in settings:
                try:
                    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,path,0,winreg.KEY_READ|winreg.KEY_WRITE)
                    existed = True
                except FileNotFoundError:
                    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER,path); existed = False
                with key:
                    try: old, kind = winreg.QueryValueEx(key,name); present = True
                    except FileNotFoundError: old = kind = None; present = False
                    saved.append(dict(path=path,name=name,key_existed=existed,present=present,value=old,kind=kind))
                    winreg.SetValueEx(key,name,0,winreg.REG_DWORD,value)
            write(output/'settings-before.json',json.dumps(saved,indent=2))
            process = subprocess.Popen([sys.executable,str(Path(__file__)),'--worker',str(folder)]+(['--smoke'] if smoke else [])+(['--baseline'] if baseline else [])+(['--benchmark'] if benchmark else []))
            deadline = time.monotonic() + 240
            while process.poll() is None:
                if (folder/'pid').exists(): office_pid = int((folder/'pid').read_text())
                if office_pid and office_pid not in original_pids:
                    def visit(hwnd, unused):
                        if win32process.GetWindowThreadProcessId(hwnd)[1] != office_pid: return
                        if not win32gui.IsWindowVisible(hwnd): return
                        if win32gui.GetClassName(hwnd) == '#32770':
                            texts, buttons = [], []
                            def child(ch, unused):
                                texts.append(win32gui.GetWindowText(ch))
                                if win32gui.GetClassName(ch) == 'Button': buttons.append(ch)
                            win32gui.EnumChildWindows(hwnd,child,None)
                            dialogs.append(' | '.join(texts))
                            print('Office dialog: '+dialogs[-1], flush=True)
                            write(output/'dialogs.json',json.dumps(dialogs,indent=2))
                            if buttons:
                                button = next((b for b in buttons if win32gui.GetWindowText(b).replace('&','') in ('End','OK')),buttons[0])
                                win32gui.PostMessage(button,win32con.BM_CLICK,0,0)
                            else: win32gui.PostMessage(hwnd,win32con.WM_CLOSE,0,0)
                        ctypes.windll.user32.ShowWindowAsync(hwnd,6)
                    win32gui.EnumWindows(visit,None)
                if dialogs:
                    diagnostic = subprocess.run([sys.executable, '-c',
                        'import win32com.client,win32process,sys; a=win32com.client.GetActiveObject("Excel.Application"); assert win32process.GetWindowThreadProcessId(a.Hwnd)[1]==int(sys.argv[1]); p=a.VBE.ActiveCodePane; print(p.CodeModule.Name,p.GetSelection()); print(p.CodeModule.Lines(p.GetSelection()[0],3))', str(office_pid)], capture_output=True, text=True, timeout=5)
                    print(diagnostic.stdout, diagnostic.stderr, flush=True)
                    write(output/'compile-error.txt',diagnostic.stdout+diagnostic.stderr)
                    raise RuntimeError('Office dialog: '+str(dialogs))
                if time.monotonic() > deadline: raise TimeoutError('Excel test exceeded 240 seconds')
                time.sleep(.05)
            result = json.loads((folder/'result.json').read_text()) if (folder/'result.json').exists() else {'status':'failed'}
            result.update(dialogs=dialogs, elapsed_seconds=time.time()-started,
                          measured_at_utc=__import__('datetime').datetime.now(__import__('datetime').timezone.utc).isoformat(),
                          baseline_source='527f172 + private ByVal compilation repair' if baseline else None,
                          imported_sources={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in folder.glob('*.cls')},
                          commit=subprocess.check_output(['git','rev-parse','HEAD'],cwd=ROOT,text=True).strip(),
                          sources={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in (ROOT/'src').glob('*.cls')})
            if dialogs: result['status'] = 'failed'
            write(output/'results.json',json.dumps(result,indent=2))
            if process.returncode or result['status'] not in ('passed','baseline'): raise RuntimeError(str(result))
        finally:
            cleanup_error = None
            try:
                if process and process.poll() is None: process.kill(); process.wait(timeout=10)
                if office_pid and office_pid not in original_pids and psutil.pid_exists(office_pid):
                    p = psutil.Process(office_pid)
                    if p.name().lower() == 'excel.exe': p.terminate(); p.wait(timeout=10)
            except Exception as exc: cleanup_error = repr(exc)
            for path in folder.glob('*.log'): shutil.copy2(path,output/path.name)
            write(output/'dialogs.json',json.dumps(dialogs,indent=2))
            for item in reversed(saved):
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER,item['path'],0,winreg.KEY_WRITE) as key:
                    if item['present']: winreg.SetValueEx(key,item['name'],0,item['kind'],item['value'])
                    else:
                        try: winreg.DeleteValue(key,item['name'])
                        except FileNotFoundError: pass
                if not item['key_existed']:
                    try: winreg.DeleteKey(winreg.HKEY_CURRENT_USER,item['path'])
                    except OSError: pass
            for item in saved:
                try:
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER,item['path']) as key:
                        value,kind = winreg.QueryValueEx(key,item['name'])
                    assert item['present'] and value == item['value'] and kind == item['kind']
                except FileNotFoundError: assert not item['present']
            write(output/'settings-restored.json',json.dumps(dict(verified=[v['name'] for v in saved],cleanup_error=cleanup_error),indent=2))
            if cleanup_error: raise RuntimeError(cleanup_error)
    print('Excel checks passed; original Office settings verified',flush=True)

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--worker',type=Path)
    parser.add_argument('--smoke',action='store_true')
    parser.add_argument('--baseline',action='store_true')
    parser.add_argument('--benchmark',action='store_true')
    parser.add_argument('--output',type=Path,default=HERE/'results/latest')
    args = parser.parse_args()
    if args.worker: worker(args.worker,args.smoke,args.baseline,args.benchmark)
    else: supervise(args.output,args.smoke,args.baseline,args.benchmark)
