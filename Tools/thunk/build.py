"""Build the relocation-free x64 VM and replace its marked VBA payload."""
import argparse
import os
from pathlib import Path
import re
import subprocess
import tempfile

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]

def build(toolchain, check=False):
    env = dict(os.environ, PATH=str(toolchain) + os.pathsep + os.environ['PATH'])
    with tempfile.TemporaryDirectory(prefix='asf-thunk-build-') as temp:
        obj, binary = Path(temp)/'vm.o', Path(temp)/'vm.bin'
        test = Path(temp)/'test.exe'
        subprocess.run([str(toolchain/'gcc.exe'), '-O2', '-Wall', '-Wextra', '-Werror', str(HERE/'test_vm.c'), '-loleaut32', '-o', str(test)], check=True, env=env)
        subprocess.run([str(test)], check=True, env=env)
        subprocess.run([str(toolchain/'gcc.exe'), '-O2', '-Wall', '-Wextra', '-Werror', '-c',
                        '-ffreestanding', '-fno-builtin', '-fno-stack-protector',
                        '-fno-jump-tables', '-fno-tree-switch-conversion',
                        '-fno-asynchronous-unwind-tables', str(HERE/'vm.c'), '-o', str(obj)], check=True, env=env)
        relocations = subprocess.check_output([str(toolchain/'objdump.exe'), '-r', str(obj)], text=True, env=env)
        if 'RELOCATION RECORDS' in relocations:
            raise RuntimeError(relocations)
        symbols = subprocess.check_output([str(toolchain/'nm.exe'), str(obj)], text=True, env=env)
        if not re.search(r'^0+ T vm$', symbols, re.M):
            raise RuntimeError('Entry must be at offset zero: ' + symbols)
        subprocess.run([str(toolchain/'objcopy.exe'), '-O', 'binary', '-j', '.text', str(obj), str(binary)], check=True, env=env)
        payload = binary.read_bytes().hex()
    target = ROOT/'src/ASF_Thunk.cls'
    source = target.read_text(encoding='utf-8')
    block = "' BEGIN GENERATED THUNK\n" + ''.join('    hexCode = hexCode & "'+payload[i:i+160]+'"\n' for i in range(0,len(payload),160)) + "' END GENERATED THUNK"
    updated, count = re.subn(r"' BEGIN GENERATED THUNK.*?' END GENERATED THUNK", lambda _: block, source, flags=re.S)
    assert count == 1
    if check:
        assert updated == source, 'Embedded thunk differs; run build.py'
    else:
        target.write_bytes(updated.replace('\n','\r\n').encode('utf-8'))
    print(f'{len(payload)//2} bytes; no relocations; ABI assertions passed; payload {"verified" if check else "updated"}')

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--toolchain', type=Path, required=True)
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    build(args.toolchain, args.check)
