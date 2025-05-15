import logging
from logging.handlers import RotatingFileHandler

# Logging configuration
LOGGING_CONFIG = {
    'level': logging.INFO,
    'handlers': [
        RotatingFileHandler('./cache-and-logs/latex_clipboard.log', maxBytes=10*1024*1024, backupCount=5),
        logging.StreamHandler()
    ]
}

# Matplotlib rcParams
RC_PARAMS = {
    'text.usetex': True,
    'text.latex.preamble': r'\usepackage{amsmath}',
    'font.family': 'serif',
    'font.serif': ['Times'],
    'axes.labelsize': 12,
    'font.size': 12,
    'legend.fontsize': 12,
    'xtick.labelsize': 10,
    'ytick.labelsize': 10,
}

