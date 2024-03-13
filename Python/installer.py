import subprocess
import pip
import os

pip.main(["install", "pyinstaller"])

command = [
    "pyinstaller", 
    'DownloadsSimplified.pyw', 
    '--onefile', 
    '--noconsole', 
    '--icon', 
    './assets/icon.ico'
]

subprocess.run(command)