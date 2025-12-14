# 运行环境
# cd C:\Users\15038\Desktop\HardWare\mm_report 
# python -m venv venv
# .\venv\Scripts\Activate.ps1
# python issues.py
# pip install requests PyGithub pyinstaller
# pyinstaller --onefile --name mm_test.exe mm_test.py

import time
from datetime import datetime, timedelta, timezone
import re
import struct
import math
import os
import sys
import scipy.interpolate as intpl 
from pyvisa import constants as pyconst
import pyvisa as visa
from pyvisa import errors
import serial
import numpy as np
from scipy.io import savemat,loadmat
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import StringVar
from tkinter import filedialog
import scipy.io as sio
import ctypes
from ctypes import wintypes


# 可能需要选择import路径
from equips_v3 import instKS_34461A, UI

def test():

    # activate the user input UI
    ui_mm_control = UI()
    ui_mm_control.generat_ui()
    ui_mm_control.mainloop()

if 1:
    test()