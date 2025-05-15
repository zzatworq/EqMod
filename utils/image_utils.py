import matplotlib.pyplot as plt
import numpy as np
from PIL import Image
import io
import logging
import subprocess
import os
import tempfile

def image_to_bytes(image):
    """Convert an image to bytes."""
    buffer = io.BytesIO()
    image.save(buffer, format='PNG')
    bytes_data = buffer.getvalue()
    logging.info(f"Image bytes length: {len(bytes_data)}")
    return bytes_data

def count_non_transparent_pixels(image):
    """Count non-transparent pixels in an image."""
    if image.mode != 'RGBA':
        image = image.convert('RGBA')
    img_array = np.array(image)
    alpha_channel = img_array[:, :, 3]
    non_transparent_count = np.sum(alpha_channel > 0)
    logging.info(f"Non-transparent pixel count: {non_transparent_count}")
    return non_transparent_count

def is_image_empty(image):
    """Check if an image is empty based on non-transparent pixels."""
    pixel_count = count_non_transparent_pixels(image)
    return pixel_count < 100

def render_latex_to_image(latex_string, text_color, font_size, dpi, mode="Matplotlib"):
    """Render a LaTeX string to an image with dynamic sizing, using specified mode."""
    logging.info(f"Processing LaTeX string: {latex_string} in mode: {mode}")
    
    if mode == "Matplotlib":
        return render_latex_matplotlib(latex_string, text_color, font_size, dpi)
    elif mode == "Standalone":
        return render_latex_standalone(latex_string, text_color, font_size, dpi)
    else:
        logging.error(f"Unknown render mode: {mode}")
        return None

def render_latex_matplotlib(latex_string, text_color, font_size, dpi):
    """Render LaTeX string using Matplotlib."""
    try:
        scaled_font_size = font_size * (dpi / 100)
        base_width = 12
        base_height = 3
        fig = plt.figure(figsize=(base_width, base_height), dpi=dpi)
        ax = fig.add_axes([0, 0, 1, 1])
        ax.set_axis_off()
        tex_string = f"${latex_string}$"
        text = ax.text(0.5, 0.5, tex_string,
                      fontsize=scaled_font_size, color=text_color,
                      horizontalalignment='center',
                      verticalalignment='center',
                      transform=ax.transAxes)
        buffer = io.BytesIO()
        plt.savefig(buffer, format='png', dpi=dpi, transparent=True,
                   bbox_inches='tight', pad_inches=0.05)
        plt.close(fig)
        buffer.seek(0)
        img = Image.open(buffer).convert("RGBA")
        bbox = img.getbbox()
        if bbox:
            left, top, right, bottom = bbox
            padding = max(5, dpi // 20)
            left = max(0, left - padding)
            top = max(0, top - padding)
            right = min(img.width, right + padding)
            bottom = min(img.height, bottom + padding)
            img = img.crop((left, top, right, bottom))
            max_width = 1800
            max_height = 600
            if img.width > max_width or img.height > max_height:
                aspect_ratio = img.width / img.height if img.height > 0 else 1
                if img.width > max_width:
                    new_width = max_width
                    new_height = int(new_width / aspect_ratio)
                else:
                    new_height = max_height
                    new_width = int(aspect_ratio * new_height)
                img = img.resize((new_width, new_height), Image.LANCZOS)
            logging.info(f"Final image size: {img.size}")
        else:
            logging.warning(f"No bounding box found for equation: {latex_string}")
            return None
        if is_image_empty(img):
            logging.warning(f"Generated empty image for equation: {latex_string}")
            return None
        return img
    except Exception as e:
        logging.error(f"Error rendering LaTeX string '{latex_string}' with Matplotlib: {e}")
        return None

def render_latex_standalone(latex_string, text_color, font_size, dpi):
    """Render LaTeX string using standalone document class and dvipng."""
    tex_template = r"""
    \documentclass[preview]{standalone}
    \usepackage{amsmath}
    \usepackage{amsfonts}
    \usepackage{xcolor}
    \begin{document}
    \fontsize{%dpt}{%dpt}\selectfont
    \color{%s}
    $%s$
    \end{document}
    """
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            tex_path = os.path.join(temp_dir, "temp.tex")
            dvi_path = os.path.join(temp_dir, "temp.dvi")
            png_path = os.path.join(temp_dir, "temp.png")
            
            # Scale font size for LaTeX (base size adjusted for DPI)
            scaled_font_size = int(font_size * (dpi / 100))
            line_skip = int(scaled_font_size * 1.2)  # Line skip typically 1.2x font size
            tex_content = tex_template % (scaled_font_size, line_skip, text_color, latex_string)
            
            # Write LaTeX file
            with open(tex_path, 'w', encoding='utf-8') as f:
                f.write(tex_content)
            
            # Compile LaTeX to DVI
            latex_cmd = ["latex", "-interaction=nonstopmode", "-output-directory", temp_dir, tex_path]
            try:
                result = subprocess.run(latex_cmd, check=True, capture_output=True, text=True)
                logging.debug(f"LaTeX command output: {result.stdout}")
            except subprocess.CalledProcessError as e:
                logging.error(f"LaTeX compilation failed: {e.stderr}")
                return None
            
            # Convert DVI to PNG
            dvipng_cmd = [
                "dvipng", "-D", str(dpi), "-T", "tight",
                "-bg", "Transparent", "-o", png_path, dvi_path
            ]
            try:
                result = subprocess.run(dvipng_cmd, check=True, capture_output=True, text=True)
                logging.debug(f"dvipng command output: {result.stdout}")
            except subprocess.CalledProcessError as e:
                logging.error(f"dvipng conversion failed: {e.stderr}")
                return None
            
            if not os.path.exists(png_path):
                logging.error("No PNG output generated by dvipng")
                return None
            
            img = Image.open(png_path).convert("RGBA")
            bbox = img.getbbox()
            if bbox:
                left, top, right, bottom = bbox
                padding = max(5, dpi // 20)
                left = max(0, left - padding)
                top = max(0, top - padding)
                right = min(img.width, right + padding)
                bottom = min(img.height, bottom + padding)
                img = img.crop((left, top, right, bottom))
                max_width = 1800
                max_height = 600
                if img.width > max_width or img.height > max_height:
                    aspect_ratio = img.width / img.height if img.height > 0 else 1
                    if img.width > max_width:
                        new_width = max_width
                        new_height = int(new_width / aspect_ratio)
                    else:
                        new_height = max_height
                        new_width = int(aspect_ratio * new_height)
                    img = img.resize((new_width, new_height), Image.LANCZOS)
                logging.info(f"Final image size: {img.size}")
            else:
                logging.warning(f"No bounding box found for equation: {latex_string}")
                return None
            
            if is_image_empty(img):
                logging.warning(f"Generated empty image for equation: {latex_string}")
                return None
            
            return img
    except Exception as e:
        logging.error(f"Error rendering LaTeX string '{latex_string}' with Standalone: {e}")
        return None