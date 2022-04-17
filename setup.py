# This Python file uses the following encoding: utf-8
import os
import sys
from cx_Freeze import setup, Executable


files = ['icon.ico', 'chromedriver.exe', 'Signa - Prodemge 1.4.0.0.crx']

target = Executable(
    script = "main.py",
    base = "Win32GUI",
    icon = "icon.ico"
    )

setup(
    name = "Marcador de Exames Teorico",
    description = "Agendador de Exames Teorico - Detran-MG",
    author = "Cristiano Leite",
    options = {'build_exe' : {'include_files' : files}},
    executables = [target]
    )
