import win32com.client
import shutil
import os

# Clear gen_py cache
gen_py_path = os.path.join(os.environ['LOCALAPPDATA'], 'Temp', 'gen_py')
if os.path.exists(gen_py_path):
    shutil.rmtree(gen_py_path, ignore_errors=True)

# Rebuild Excel COM binding
excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
excel.Quit()