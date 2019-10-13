# simpchin_detect
A simple script to check if a file has Chinese text and whether it is Simplified, Traditional, or a mixture. 

Wiki needs to be updated.

Compilation notes (these steps necessary to ensure that package is small; otherwise, 200+ MB size)
1. Created virtual environment. Deactivate base anaconda environment
2. In virtual env, installed all libraries (zhon, xlrd, python-pptx, pypiwin32)
3. In virtual env, installed pyinstaller
4. Using pyinstaller -w -F (meaning not windowed and onefile), compiled script