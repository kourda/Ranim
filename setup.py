from doctest import debug_script
from ssl import Options
import sys
import os
from cx_Freeze import setup,Executable
files=['icon.ico','init.ui','second.ui','stats3G.xls','stats2G.xls','stats4G.xls']
target = Executable(
	script="main.py",
	icon="icon.ico",
	base="Win32GUI"
)

setup(
	name = "rapport de maintenance",
	version = "1.0",
	description = "Seuills",
	author = "Ranim",
	options = {"build_exe":{'include_files':files}},
	executables = [target]
)
