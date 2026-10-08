"""用一次ODA进程批量转换DWG副本，保留逐文件回退入口。"""
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from app.core.config import settings
from app.service.cad_text_service import find_oda_file_converter, _headless_command

request=json.loads(Path(sys.argv[1]).read_text(encoding='utf-8'))
base=Path(request['output_dir'])/'CAD批量转换'
source_dir=base/'input'
target_dir=base/'output'
source_dir.mkdir(parents=True,exist_ok=True)
target_dir.mkdir(exist_ok=True)
manifest=json.loads(Path(request['manifest']).read_text(encoding='utf-8'))
unique={r['sha256']:r for r in manifest if r['extension']=='.dwg'}
mapping={}
for index,(digest,row) in enumerate(unique.items(),1):
    name=f'{index:03d}'
    shutil.copyfile(Path(request['root'])/row['relative_path'],source_dir/(name+'.dwg'))
    mapping[digest]=str(target_dir/(name+'.dxf'))
converter=find_oda_file_converter(settings.ODA_FILE_CONVERTER_PATH)
if converter:
    env=os.environ.copy()
    command=_headless_command([str(converter),str(source_dir),str(target_dir),'ACAD2018','DXF','0','1','*.dwg'],env)
    startupinfo=subprocess.STARTUPINFO()
    startupinfo.dwFlags|=subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow=subprocess.SW_HIDE
    with (base/'ODA运行日志.txt').open('wb') as log:
        process=subprocess.Popen(command,env=env,stdout=log,stderr=log,startupinfo=startupinfo,creationflags=subprocess.CREATE_NO_WINDOW)
        started=time.monotonic()
        while process.poll() is None:
            if time.monotonic()-started>3600:
                process.kill()
                break
            print(json.dumps({'cad_done':len(list(target_dir.glob('*.dxf'))),'cad_total':len(unique)}),flush=True)
            time.sleep(10)
        process.wait()
mapping={k:v for k,v in mapping.items() if Path(v).is_file()}
(base/'cad_map.json').write_text(json.dumps(mapping),encoding='utf-8')
print(json.dumps({'cad_done':len(mapping),'cad_total':len(unique)}),flush=True)
