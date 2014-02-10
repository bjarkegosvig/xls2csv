from cx_Freeze import setup, Executable
# run as python cx_build build
exe = Executable(
      script="gui.py",
      base="Win32GUI",
      targetName="xls2csv.exe",
      icon = "excel.ico"

     )

setup( name = "xls2csv",
           version = "0.1",
           description = "Convert xls file to csv",
           executables = [exe]
         )