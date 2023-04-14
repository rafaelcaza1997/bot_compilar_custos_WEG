import sys
from cx_Freeze import setup, Executable

# base = None
# if sys.platform == "win32":
#     base = "Win32GUI"
    

build_exe_options = {
                        "packages": ["os","pandas","datetime","pathlib","win32com.client","configparser","openpyxl","fsspec","xlsxwriter","requests","json"],
                        "build_exe" : r'C:\Users\rafae\OneDrive\Projetos\Python\WEG\Bot Custos\exe\v6',
                    }

                        # "excludes": ["tkinter",'sqlite3'],

exe = Executable(
    script="main.py",
    base = "Win32GUI",
                          # Retirar comentario se for contruir um executavel para windows
    )
setup(
    name = "bot_custos",
    version = "1.0.0",
    
    options = {'build_exe': build_exe_options},
    executables = [exe],
    )
# python setup.py bdist_msi














