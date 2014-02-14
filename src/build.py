from cx_Freeze import setup, Executable
# run as python cx_build build
includefiles = ['..\UNLICENSE.txt', '..\README.md']
includes = ['csvwriter','formatxls']
excludes = []
packages = ['win32com.gen_py']

exe = Executable(
      script="gui.py",
      base="Win32GUI",
      targetName="xls2csv.exe",
      #targetDir = r"dist",
      #compress = True,
      copyDependentFiles = True,
      #appendScriptToExe = True,
      #appendScriptToLibrary = True,
      icon = "excel.ico"

     )

setup( name = "xls2csv",
           version = "0.11",
           description = "Convert xls file to csv",
          options = {'build_exe': {'excludes':excludes,'packages':packages,'includes':includes,'include_files':includefiles}},
           executables = [exe]
         )