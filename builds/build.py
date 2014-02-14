from cx_Freeze import setup, Executable
# run as python cx_build build
includes = []
excludes = []
packages = ['win32com.gen_py']

exe = Executable(
      script="../source/gui.py",
      base="Win32GUI",
      targetName="xls2csv.exe",
      icon = "excel.ico"

     )

setup( name = "xls2csv",
           version = "0.11",
           description = "Convert xls file to csv",
          options = {'build_exe': {'excludes':excludes,'packages':packages,'includes':includes}},
           executables = [exe]
         )