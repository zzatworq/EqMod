import os
import subprocess

# Add MiKTeX LaTeX path to Python's PATH environment variable
os.environ['PATH'] += os.pathsep + r'C:\Program Files\MiKTeX\miktex\bin\x64'  # Adjust path if necessary

# Now check if LaTeX is accessible
try:
    result = subprocess.run(['latex', '--version'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    print(result.stdout.decode())
except FileNotFoundError:
    print("LaTeX executable not found.")
