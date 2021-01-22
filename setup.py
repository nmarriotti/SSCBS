from cx_Freeze import setup, Executable

base = None    

executables = [Executable("get_packages.py", base=base)]

packages = ["idna","paramiko","time","openpyxl","sys","os"]
options = {
    'build_exe': {    
        'packages':packages,
    },    
}

setup(
    name = "SSCBS",
    options = options,
    version = "1.3.1",
    description = 'Connects to Linux hosts and populates Excel spreadsheet with installed packages.',
    executables = executables
)
