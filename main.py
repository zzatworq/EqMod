import ctypes
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from matplotlib import rcParams
import tkinter as tk
import logging
from configs.settings import LOGGING_CONFIG, RC_PARAMS
from gui.app_gui import LatexClipboardApp

# Configure logging
for handler in LOGGING_CONFIG['handlers']:
    logging.getLogger().addHandler(handler)
logging.getLogger().setLevel(LOGGING_CONFIG['level'])

# Update matplotlib parameters
rcParams.update(RC_PARAMS)

# Enable DPI awareness
try:
    ctypes.windll.user32.SetProcessDPIAware()
except Exception:
    pass

if __name__ == "__main__":
    root = tk.Tk()
    app = LatexClipboardApp(root)
    root.mainloop()