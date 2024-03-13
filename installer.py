import subprocess

command = [
    ".venv\Scripts\pyinstaller.exe", 
    'DownloadsSimplified.pyw', 
    '--onefile', 
    '--noconsole', 
    '--icon', 
    './Images/icon.ico'
]

subprocess.run(command)
