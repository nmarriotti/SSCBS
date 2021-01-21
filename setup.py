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
    name = "RPM2Excel",
    options = options,
    version = "1.0.0",
    description = 'Connects to Linux hosts and populates Excel spreadsheet with installed packages.',
    executables = executables
)
