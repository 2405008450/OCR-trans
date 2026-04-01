import os
import glob

html_files = glob.glob('e:/fastapi-llm-demo/static/*.html')
script_tag = '<script src="/static/header.js"></script>'

for file_path in html_files:
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    if script_tag not in content:
        content = content.replace('</body>', f'    {script_tag}\n</body>')
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f'Injected to {os.path.basename(file_path)}')

print('Injection script completed.')
