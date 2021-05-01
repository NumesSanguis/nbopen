"""Install GUI integration on Windows"""

import sys
import os

# TODO as argument
style = "jupyter"  # voila
template = ""  # --template vuetify-default   # with space at the end if not empty

try:
  import winreg
except ImportError:  
  import _winreg as winreg

SZ = winreg.REG_SZ
with winreg.CreateKey(winreg.HKEY_CURRENT_USER, "Software\Classes\.ipynb") as k:
    winreg.SetValue(k, "", SZ, "Jupyter.nbopen")
    winreg.SetValueEx(k, "Content Type", 0, SZ, "application/x-ipynb+json")
    winreg.SetValueEx(k, "PerceivedType", 0, SZ, "document")
    with winreg.CreateKey(k, "OpenWithProgIds") as openwith:
        winreg.SetValueEx(openwith, "Jupyter.nbopen", 0, winreg.REG_NONE, b'')

# check if we're in a conda env
executable = sys.executable
try:
    conda_env = os.environ['CONDA_DEFAULT_ENV']
    # TODO automatically find Anaconda python.exe (Admin install ProgramData, otherwise
    #  C:\Users\User-Name\Anaconda3\Scripts\anaconda.exe)
    launch_cmd = f'"C:\ProgramData\Anaconda3\python.exe" -m conda run -n {conda_env} pythonw -m '
    if style == "jupyter":
        launch_cmd += 'nbopen "%1"'
    elif style == "voila":
        launch_cmd += f'voila {template}"%1"'
    else:
        raise ValueError(f"style '{style}' not supported.")
    
    print(f"Anaconda environment found: {conda_env}")
    print(f"Setting up command:\n{launch_cmd}")
# TODO check branches for new commands
except KeyError as e:
    print(f"Install script not called in a Conda environment:\n{e}")

    if executable.endswith("python.exe"):
        executable = executable[:-10] + 'pythonw.exe'
    launch_cmd = f'"{executable}" -m nbopen "%1"'
    
    print(f"Setting up old command (not working in Windows 10)\n{launch_cmd}")

with winreg.CreateKey(winreg.HKEY_CURRENT_USER, "Software\Classes\Jupyter.nbopen") as k:
    winreg.SetValue(k, "", SZ, "IPython notebook")
    with winreg.CreateKey(k, "shell\open\command") as launchk:
        winreg.SetValue(launchk, "", SZ, launch_cmd)

try:
    from win32com.shell import shell, shellcon
    shell.SHChangeNotify(shellcon.SHCNE_ASSOCCHANGED, shellcon.SHCNF_IDLIST, None, None)
except ImportError:
    print("You may need to restart for association with .ipynb files to work")
    print("  (pywin32 is needed to notify Windows of the change)")
