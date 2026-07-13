import re

with open('main.py', 'r', encoding='utf-8') as f:
    content = f.read()

old_zhengwen = r'r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查1\llm\llm_project\zhengwen\output_excel"'
old_yemei    = r'r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查1\llm\llm_project\yemei\output_excel"'
old_yejiao   = r'r"D:\Users\Administrator\Desktop\项目文件\专检\数值检查1\llm\llm_project\yejiao\output_excel"'

content = content.replace(old_zhengwen, 'str(LLM_DIR / "zhengwen" / "output_excel")')
content = content.replace(old_yemei,    'str(LLM_DIR / "yemei" / "output_excel")')
content = content.replace(old_yejiao,   'str(LLM_DIR / "yejiao" / "output_excel")')

with open('main.py', 'w', encoding='utf-8') as f:
    f.write(content)

print('done')
