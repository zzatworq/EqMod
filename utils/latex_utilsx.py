import subprocess
import os
import tempfile
import logging
import re
from PIL import Image
from templates.test_string import TEST_STRING
text = TEST_STRING
def check_latex():
    """
    Check if the input text contains valid LaTeX equations.
    
    Args:
        text (str): Input text to check.
    
    Returns:
        bool: True if LaTeX equations are found, False otherwise.
    """
    if not isinstance(text, str):
        logging.error("Input to check_latex must be a string")
        return False
    
    # Patterns for inline ($...$) and display ($$...$$, \[...\], \begin{equation}...\end{equation})
    latex_patterns = [
        r'\$[^\$]+\$',  # Inline: $...$
        r'\\\[[^\\]+\]',  # Display: \[...\]
        r'\\begin\{equation\}[^\\]+\\end\{equation\}',  # Display: \begin{equation}...\end{equation}
        r'\$\$[^\$]+\$\$'  # Display: $$...$$
    ]
    
    for pattern in latex_patterns:
        if re.search(pattern, text, re.MULTILINE | re.DOTALL):
            return True
    return False

def find_latex_equations(text):
    """
    Extract LaTeX equations from text, preserving order and type (inline/display).
    
    Args:
        text (str): Input text containing LaTeX equations.
    
    Returns:
        list: List of tuples (equation, is_inline), where equation is the LaTeX string
              and is_inline is True for inline equations, False for display.
    """
    equations = []
    
    # Inline: $...$ (non-greedy match)
    inline_pattern = r'\$(.+?)\$'
    for match in re.finditer(inline_pattern, text, re.MULTILINE | re.DOTALL):
        eq = match.group(1).strip()
        if eq:
            equations.append((eq, True))
    
    # Display: \[...\]
    display_pattern1 = r'\\\[(.+?)\\\]'
    for match in re.finditer(display_pattern1, text, re.MULTILINE | re.DOTALL):
        eq = match.group(1).strip()
        if eq:
            equations.append((eq, False))
    
    # Display: \begin{equation}...\end{equation}
    display_pattern2 = r'\\begin\{equation\}(.+?)\\end\{equation\}'
    for match in re.finditer(display_pattern2, text, re.MULTILINE | re.DOTALL):
        eq = match.group(1).strip()
        if eq:
            equations.append((eq, False))
    
    # Display: $$...$$
    display_pattern3 = r'\$\$(.+?)\$\$'
    for match in re.finditer(display_pattern3, text, re.MULTILINE | re.DOTALL):
        eq = match.group(1).strip()
        if eq:
            equations.append((eq, False))
    
    logging.info(f"Found {len(equations)} LaTeX equations")
    return equations

def render_latex_to_image(latex_str, dpi=300):
    """
    Render a LaTeX equation to a PIL Image using luatex, suppressing console popups.
    
    Args:
        latex_str (str): LaTeX equation string (e.g., 'E=mc^{2^{2}}').
        dpi (int): Resolution for the output image.
    
    Returns:
        PIL.Image or None: Rendered image or None if rendering fails.
    """
    # LaTeX template for standalone equation
    tex_template = r"""
    \documentclass[preview]{standalone}
    \usepackage{amsmath}
    \usepackage{amsfonts}
    \begin{document}
    $%s$
    \end{document}
    """

    # Create temporary directory and files
    with tempfile.TemporaryDirectory() as tmpdir:
        tex_file = os.path.join(tmpdir, "temp.tex")
        png_file = os.path.join(tmpdir, "temp.png")
        
        # Write LaTeX file
        with open(tex_file, "w") as f:
            f.write(tex_template % latex_str)
        
        try:
            # Run luatex to generate PNG, suppressing console
            subprocess.run(
                [
                    "luatex",
                    "--output-format=png",
                    f"--resolution={dpi}",
                    f"--output-directory={tmpdir}",
                    tex_file
                ],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                creationflags=0x08000000  # CREATE_NO_WINDOW
            )
            
            # Check if PNG was generated
            if not os.path.exists(png_file):
                logging.warning(f"No image generated for equation: {latex_str}")
                return None
            
            # Open and return PIL Image
            with Image.open(png_file) as img:
                return img.convert("RGBA")
                
        except subprocess.CalledProcessError as e:
            logging.error(f"luatex failed for '{latex_str}': {e.stderr.decode()}")
            return None
        except FileNotFoundError:
            logging.error("luatex executable not found. Ensure MiKTeX is installed.")
            return None