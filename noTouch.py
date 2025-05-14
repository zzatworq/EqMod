import matplotlib.pyplot as plt
from matplotlib import rcParams
import numpy as np
from PIL import Image
import io
import win32clipboard
import win32con
import time
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import subprocess
import sys
import logging
import os
import base64
from docx import Document
from docx.shared import Pt

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('latex_monitor.log'),
        logging.StreamHandler()
    ]
)

rcParams['text.usetex'] = True
rcParams['text.latex.preamble'] = r'\usepackage{amsmath}'

def check_latex():
    try:
        subprocess.run(['latex', '--version'], capture_output=True, check=True)
        logging.info("LaTeX distribution found")
        return True
    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        logging.error(f"LaTeX not found: {e}")
        return False

def validate_base64(b64_data):
    try:
        decoded = base64.b64decode(b64_data, validate=True)
        if not decoded.startswith(b'\x89PNG\r\n\x1a\n'):
            logging.error("Base64 data is not a valid PNG")
            return False
        return True
    except Exception as e:
        logging.error(f"Invalid base64 data: {e}")
        return False

def set_clipboard_html(html_content):
    html_content = html_content.replace('\n', '\r\n')
    
    html = f"""
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
</HEAD>
<BODY>
<!--StartFragment-->
{html_content}
<!--EndFragment-->
</BODY>
</HTML>
"""
    
    byte_html = html.encode('utf-8')
    start_fragment = byte_html.find(b'<!--StartFragment-->') + len(b'<!--StartFragment-->')
    end_fragment = byte_html.find(b'<!--EndFragment-->')
    
    header = f"""Version:0.9
StartHTML:00000000
EndHTML:{len(byte_html):08d}
StartFragment:{start_fragment:08d}
EndFragment:{end_fragment:08d}
"""
    
    header_bytes = header.encode('utf-8')
    start_html = len(header_bytes)
    
    header = f"""Version:0.9
StartHTML:{start_html:08d}
EndHTML:{len(header_bytes) + len(byte_html):08d}
StartFragment:{start_html + start_fragment:08d}
EndFragment:{start_html + end_fragment:08d}
"""
    
    clipboard_data = header.encode('utf-8') + byte_html
    
    logging.info(f"HTML clipboard data size: {len(clipboard_data)} bytes")
    img_count = len(re.findall(r'<img src="data:image/png;base64,', html))
    logging.info(f"HTML content contains {img_count} image tags")
    
    try:
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
            win32clipboard.SetClipboardData(cf_html, clipboard_data)
            plain_text = "LaTeX Equations and Text (Images and Text - Paste in Office Application)"
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, plain_text)
            logging.info("HTML clipboard data set successfully")
        finally:
            win32clipboard.CloseClipboard()
    except Exception as e:
        logging.error(f"Failed to set clipboard data: {e}")
        raise

class LatexClipboardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("LaTeX Clipboard Monitor")
        self.monitoring = False
        self.monitor_thread = None
        self.stop_event = threading.Event()
        self.last_images = []
        self.last_text = ""

        if not check_latex():
            messagebox.showerror("LaTeX Not Found",
                                 "A LaTeX distribution (e.g., MiKTeX) is required. Please install MiKTeX and try again.")
            sys.exit(1)

        self.create_gui()
        logging.info("Using white text on transparent background. Ensure document background is non-white for visibility.")

    def create_gui(self):
        tk.Label(self.root, text="Text Color:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.color_var = tk.StringVar(value="white")
        colors = ["white", "black", "red", "blue", "green"]
        self.color_menu = ttk.OptionMenu(self.root, self.color_var, "white", *colors)
        self.color_menu.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        tk.Label(self.root, text="Font Size:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.font_size_var = tk.StringVar(value="12")
        self.font_size_spin = ttk.Spinbox(self.root, from_=10, to=50, width=10, textvariable=self.font_size_var)
        self.font_size_spin.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        tk.Label(self.root, text="DPI:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.dpi_var = tk.StringVar(value="300")
        self.dpi_spin = ttk.Spinbox(self.root, from_=100, to=600, width=10, textvariable=self.dpi_var)
        self.dpi_spin.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        self.only_images_var = tk.BooleanVar(value=False)
        self.only_images_check = ttk.Checkbutton(self.root, text="Only Images (Ignore Text)", variable=self.only_images_var)
        self.only_images_check.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky="w")

        self.toggle_button = ttk.Button(self.root, text="Start Monitoring", command=self.toggle_monitoring)
        self.toggle_button.grid(row=4, column=0, pady=10)

        self.test_button = ttk.Button(self.root, text="Test Render", command=self.test_render)
        self.test_button.grid(row=4, column=1, pady=10)

        self.save_button = ttk.Button(self.root, text="Save as DOCX", command=self.save_as_docx)
        self.save_button.grid(row=4, column=2, pady=10)

        tk.Label(self.root, text="Note: White text may be invisible on white backgrounds.", fg="red").grid(row=5, column=0, columnspan=3, padx=5, pady=5)

        self.status_var = tk.StringVar(value="Stopped")
        tk.Label(self.root, textvariable=self.status_var, fg="red").grid(row=6, column=0, columnspan=3, pady=5)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def validate_inputs(self):
        try:
            font_size = int(self.font_size_var.get())
            if not 10 <= font_size <= 50:
                raise ValueError("Font size must be between 10 and 50")
            dpi = int(self.dpi_var.get())
            if not 100 <= dpi <= 600:
                raise ValueError("DPI must be between 100 and 600")
            return True
        except ValueError as e:
            logging.error(f"Input validation failed: {e}")
            messagebox.showerror("Invalid Input", str(e))
            return False

    def toggle_monitoring(self):
        if not self.monitoring:
            if not self.validate_inputs():
                return
            self.monitoring = True
            self.status_var.set("Monitoring")
            self.toggle_button.configure(text="Stop Monitoring")
            self.color_menu.configure(state="disabled")
            self.font_size_spin.configure(state="disabled")
            self.dpi_spin.configure(state="disabled")
            self.test_button.configure(state="disabled")
            self.save_button.configure(state="disabled")
            self.stop_event.clear()
            self.monitor_thread = threading.Thread(target=self.monitor_clipboard, daemon=True)
            self.monitor_thread.start()
            logging.info("Started clipboard monitoring")
        else:
            self.monitoring = False
            self.status_var.set("Stopped")
            self.toggle_button.configure(text="Start Monitoring")
            self.color_menu.configure(state="normal")
            self.font_size_spin.configure(state="normal")
            self.dpi_spin.configure(state="normal")
            self.test_button.configure(state="normal")
            self.save_button.configure(state="normal")
            self.stop_event.set()
            logging.info("Stopped clipboard monitoring")

    def test_render(self):
        if not self.validate_inputs():
            return
        test_text = r'''To simplify the expression:

\[
\left( \frac{20}{x^2 - 36} - \frac{2}{x - 6} \right) \times \frac{1}{4 - x}
\]

we follow these steps:

### Step 1: Factor the Denominators
First, factor the denominators where possible.

\[
x^2 - 36 = (x - 6)(x + 6)
\]

So, the expression becomes:

\[
\left( \frac{20}{(x - 6)(x + 6)} - \frac{2}{x - 6} \right) \times \frac{1}{4 - x}
\]

### Step 2: Combine the Fractions Inside the Parentheses
To combine the fractions, find a common denominator, which is \((x - 6)(x + 6)\).

\[
\frac{20}{(x - 6)(x + 6)} - \frac{2}{x - 6} = \frac{20 - 2(x + 6)}{(x - 6)(x + 6)}
\]

Simplify the numerator:

\[
20 - 2(x + 6) = 20 - 2x - 12 = 8 - 2x
\]

So, the combined fraction is:

\[
\frac{8 - 2x}{(x - 6)(x + 6)}
\]

### Step 3: Factor the Numerator
Factor out a common term from the numerator:

\[
8 - 2x = 2(4 - x)
\]

Now, the expression becomes:

\[
\frac{2(4 - x)}{(x - 6)(x + 6)} \times \frac{1}{4 - x}
\]

### Step 4: Simplify the Expression
Notice that \((4 - x)\) appears in both the numerator and the denominator, so they cancel out:

\[
\frac{2}{(x - 6)(x + 6)}
\]

### Final Answer
The simplified form of the expression is:

\[
\boxed{\frac{2}{(x - 6)(x + 6)}}
\] Test equation: \[E=mc^{2^{2}}\] and another \( \int_0^1 x^2 dx = \frac{1}{3} \). Normal text follows: \[\sum M_A = 0\].'''
        logging.info(f"Rendering test text: {test_text}")
        try:
            equations = self.find_latex_equations(test_text)
            images = []
            for eq in equations['equations']:
                img = self.render_latex_to_image(eq)
                if img:
                    images.append(img)
                    debug_path = f"test_eq_{len(images)-1}.png"
                    img.save(debug_path, format='PNG')
                    logging.info(f"Saved test debug image: {debug_path}")
                else:
                    logging.warning(f"Failed to render equation: {eq}")
            
            if images:
                self.copy_images(images, test_mode=True, original_text=test_text, equations=equations)
                self.status_var.set(f"Copied {len(images)} images with text")
                messagebox.showinfo("Test Render", f"Copied {len(images)} images with text as HTML")
            else:
                self.status_var.set("No valid test images rendered")
                messagebox.showerror("Test Render", "Failed to render any valid test images")
        except Exception as e:
            logging.error(f"Test render failed: {e}")
            messagebox.showerror("Test Render Failed", f"Error rendering test equations: {e}")
            self.status_var.set("Test render failed")

    def save_as_docx(self):
        if not self.last_images:
            messagebox.showwarning("No Images", "No images available to save. Run Test Render or monitor clipboard first.")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Save as DOCX",
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            doc = Document()
            for img in self.last_images:
                img_path = f"temp_eq.png"
                img.save(img_path, format='PNG')
                doc.add_picture(img_path, width=Pt(300))
                os.remove(img_path)
            
            doc.save(file_path)
            messagebox.showinfo("Save Successful", f"Saved {len(self.last_images)} equations to {file_path}")
        except Exception as e:
            logging.error(f"Failed to save DOCX: {e}")
            messagebox.showerror("Save Failed", f"Error saving DOCX file: {e}")

    def on_closing(self):
        if self.monitoring:
            self.stop_event.set()
            if self.monitor_thread:
                self.monitor_thread.join(timeout=1.0)
        self.root.destroy()
        logging.info("Application closed")

    def image_to_bytes(self, image):
        buffer = io.BytesIO()
        image.save(buffer, format='PNG')
        bytes_data = buffer.getvalue()
        logging.info(f"Image bytes length: {len(bytes_data)}")
        return bytes_data

    def count_non_transparent_pixels(self, image):
        if image.mode != 'RGBA':
            image = image.convert('RGBA')
        img_array = np.array(image)
        alpha_channel = img_array[:, :, 3]
        non_transparent_count = np.sum(alpha_channel > 0)
        logging.info(f"Non-transparent pixel count: {non_transparent_count}")
        return non_transparent_count

    def is_image_empty(self, image):
        pixel_count = self.count_non_transparent_pixels(image)
        return pixel_count < 100

    def copy_images(self, images, test_mode=False, original_text="", equations=None):
        if not images:
            logging.info("No images to copy")
            return
        
        self.last_images = []
        html_content = ""
        text_color = self.color_var.get()
        font_size = int(self.font_size_var.get())
        
        style = f"""
        <style>
            body {{
                color: {text_color};
                font-family: Arial, sans-serif;
                font-size: {font_size}pt;
                line-height: 1.5;
            }}
            p, div, span {{
                color: {text_color} !important;
            }}
            img {{
                vertical-align: middle;
                margin: 2px 0;
            }}
        </style>
        """
        
        if test_mode or original_text:
            if not equations:
                logging.warning("No equations provided for text replacement")
                return
            
            valid_images = []
            valid_b64_data = []
            for i, img in enumerate(images):
                if img is None or self.is_image_empty(img):
                    logging.warning(f"Skipping empty/null image at index {i}")
                    continue
                
                try:
                    if img.width == 0 or img.height == 0:
                        logging.warning(f"Skipping zero-dimension image at index {i}")
                        continue
                    
                    bytes_data = self.image_to_bytes(img)
                    b64_data = base64.b64encode(bytes_data).decode('ascii')
                    
                    if not validate_base64(b64_data):
                        logging.warning(f"Skipping invalid base64 data for image {i}")
                        continue
                    
                    valid_images.append(img)
                    valid_b64_data.append(b64_data)
                    self.last_images.append(img)
                    
                    debug_path = f"debug_eq_{i}.png"
                    img.save(debug_path, format='PNG')
                    logging.info(f"Saved debug image: {debug_path}")
                
                except Exception as e:
                    logging.error(f"Error processing image {i}: {e}")
                    continue
            
            if not valid_images:
                logging.warning("No valid images to copy to clipboard")
                self.status_var.set("No valid images")
                return
            
            html_content += style
            
            if self.only_images_var.get():
                for b64_data in valid_b64_data:
                    html_content += f'<img src="data:image/png;base64,{b64_data}" style="vertical-align: middle; margin: 2px 0;">'
            else:
                last_pos = 0
                img_index = 0
                for match in equations['matches']:
                    start, end = match['start'], match['end']
                    text_segment = original_text[last_pos:start]
                    text_segment = (text_segment.replace('&', '&amp;')
                                  .replace('<', '&lt;')
                                  .replace('>', '&gt;')
                                  .replace('\r\n', '<br>')
                                  .replace('\n', '<br>'))
                    html_content += f'<span>{text_segment}</span>'
                    
                    if img_index < len(valid_b64_data):
                        html_content += f'<img src="data:image/png;base64,{valid_b64_data[img_index]}" style="vertical-align: middle; margin: 2px 0;">'
                        img_index += 1
                    
                    last_pos = end
                
                text_segment = original_text[last_pos:]
                text_segment = (text_segment.replace('&', '&amp;')
                              .replace('<', '&lt;')
                              .replace('>', '&gt;')
                              .replace('\r\n', '<br>')
                              .replace('\n', '<br>'))
                html_content += f'<span>{text_segment}</span>'
        else:
            html_content += style
            for i, img in enumerate(images):
                if img is None or self.is_image_empty(img):
                    logging.warning(f"Skipping empty/null image at index {i}")
                    continue
                
                try:
                    if img.width == 0 or img.height == 0:
                        logging.warning(f"Skipping zero-dimension image at index {i}")
                        continue
                    
                    bytes_data = self.image_to_bytes(img)
                    b64_data = base64.b64encode(bytes_data).decode('ascii')
                    
                    if not validate_base64(b64_data):
                        logging.warning(f"Skipping invalid base64 data for image {i}")
                        continue
                    
                    html_content += f'<img src="data:image/png;base64,{b64_data}" style="vertical-align: middle; margin: 2px 0;">'
                    if i < len(images) - 1:
                        html_content += '<br>'
                    
                    self.last_images.append(img)
                    
                    debug_path = f"debug_eq_{i}.png"
                    img.save(debug_path, format='PNG')
                    logging.info(f"Saved debug image: {debug_path}")
                
                except Exception as e:
                    logging.error(f"Error processing image {i}: {e}")
                    continue
        
        if html_content and self.last_images:
            try:
                set_clipboard_html(html_content)
                self.status_var.set(f"Copied {len(self.last_images)} images with text")
                logging.info(f"Successfully copied {len(self.last_images)} images with text to clipboard")
            except Exception as e:
                logging.error(f"Failed to set clipboard data: {e}")
                self.status_var.set("Error copying images")
        else:
            logging.warning("No valid images to copy to clipboard")
            self.status_var.set("No valid images")

    def render_latex_to_image(self, latex_string):
        logging.info(f"Processing LaTeX string: {latex_string}")
        try:
            text_color = self.color_var.get()
            font_size = int(self.font_size_var.get())
            dpi = int(self.dpi_var.get())

            fig = plt.figure(figsize=(3, 1), dpi=dpi)
            ax = fig.add_axes([0, 0, 1, 1])
            ax.set_axis_off()
            
            tex_string = f"${latex_string}$"
            ax.text(0, 0, tex_string,
                    fontsize=font_size, color=text_color,
                    verticalalignment='center',
                    transform=ax.transAxes)
            
            buffer1 = io.BytesIO()
            plt.savefig(buffer1, format='png', dpi=dpi, transparent=True, pad_inches=0.1)
            plt.close(fig)
            
            buffer1.seek(0)
            img1 = Image.open(buffer1).convert("RGBA")
            bbox = img1.getbbox()
            
            if bbox:
                left, top, right, bottom = bbox
                width, height = right - left, bottom - top
                
                max_width = 1800
                scale = min(1.0, max_width / width) if width > max_width else 1.0
                scaled_width = width * scale
                scaled_height = height * scale
                
                padding = dpi // 10
                fig_width = max((scaled_width + 2 * padding) / dpi, 0.3)
                fig_height = max((scaled_height + 2 * padding) / dpi, 0.1)
                logging.info(f"Scaled dimensions: width={scaled_width}, height={scaled_height}")
            else:
                logging.warning(f"No content found for equation: {latex_string}")
                fig_width, fig_height = 0.3, 0.1
            
            fig = plt.figure(figsize=(fig_width, fig_height), dpi=dpi)
            ax = fig.add_axes([0, 0, 1, 1])
            ax.set_axis_off()
            
            ax.text(0.5, 0.5, tex_string,
                    fontsize=font_size, color=text_color,
                    horizontalalignment='center',
                    verticalalignment='center',
                    transform=ax.transAxes)
            
            buffer2 = io.BytesIO()
            plt.savefig(buffer2, format='png', dpi=dpi, transparent=True, pad_inches=0.05)
            plt.close(fig)
            
            buffer2.seek(0)
            final_img = Image.open(buffer2).convert("RGBA")
            
            bbox = final_img.getbbox()
            if bbox:
                left, top, right, bottom = bbox
                padding = max(5, dpi // 30)
                
                img_width, img_height = final_img.size
                left = max(0, left - padding)
                top = max(0, top - padding)
                right = min(img_width, right + padding)
                bottom = min(img_height, bottom + padding)
                
                final_img = final_img.crop((left, top, right, bottom))
                
                if final_img.width > 1800:
                    ratio = 1800 / final_img.width
                    new_height = int(final_img.height * ratio)
                    final_img = final_img.resize((1800, new_height), Image.LANCZOS)
            else:
                logging.warning(f"No valid bounding box for equation: {latex_string}")
                return None
            
            if self.is_image_empty(final_img):
                logging.warning(f"Generated empty image for equation: {latex_string}")
                return None
            
            logging.info(f"Final image size: {final_img.size}")
            return final_img
        except Exception as e:
            logging.error(f"Error rendering LaTeX string '{latex_string}': {e}")
            return None

    def get_clipboard_text(self):
        try:
            win32clipboard.OpenClipboard()
            try:
                if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                    text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
                    if text:
                        logging.info(f"Retrieved clipboard text: {text[:100]}...")
                    return text
                else:
                    logging.info("Clipboard contains non-text data")
                    return None
            finally:
                win32clipboard.CloseClipboard()
        except win32clipboard.error as e:
            logging.error(f"Error accessing clipboard: {e}")
            return None

    def find_latex_equations(self, text):
        if not text:
            logging.info("No text provided for equation detection")
            return {'equations': [], 'matches': []}
        
        patterns = [
            (r'\\\[(.*?)\\\]', r'\\\[(.*?[^\\])\\\]'),
            (r'\\\((.*?)\\\)', r'\\\((.*?[^\\])\\\)'),
            (r'\$\$(.*?)\$\$', r'\$\$(.*?[^$])\$\$'),
            (r'\$(.*?)\$', r'\$(.*?[^$])\$'),
            (r'\\begin\{equation\}(.*?)\\end\{equation\}', r'\\begin\{equation\}(.*?[^\\])\\end\{equation\}'),
        ]
        
        equations = []
        matches = []
        text_copy = text
        
        for raw_pattern, pattern in patterns:
            for match in re.finditer(pattern, text, re.DOTALL):
                equation = match.group(1).strip()
                if equation and not re.match(r'^\s*$', equation) and not re.search(r'\.\*\?', equation) and equation != '\\':
                    equations.append(equation)
                    matches.append({
                        'start': match.start(),
                        'end': match.end(),
                        'full_match': match.group(0)
                    })
        
        matches.sort(key=lambda x: x['start'])
        
        logging.info(f"Found {len(equations)} valid equations in text")
        for i, eq in enumerate(equations):
            logging.info(f"Equation {i+1}: {eq}")
        return {'equations': equations, 'matches': matches}

    def monitor_clipboard(self):
        last_sequence = None
        processed_text = None
        
        while not self.stop_event.is_set():
            try:
                current_sequence = win32clipboard.GetClipboardSequenceNumber()
                if current_sequence != last_sequence:
                    last_sequence = current_sequence
                    text = self.get_clipboard_text()
                    
                    if text and text != processed_text:
                        logging.info("New clipboard content detected")
                        equations = self.find_latex_equations(text)
                        
                        if equations['equations']:
                            logging.info(f"Processing {len(equations['equations'])} equation(s)")
                            images = []
                            for i, eq in enumerate(equations['equations'], 1):
                                logging.info(f"Processing equation {i}: {eq}")
                                try:
                                    img = self.render_latex_to_image(eq)
                                    if img:
                                        images.append(img)
                                    else:
                                        logging.warning(f"Equation {i} produced no valid image")
                                except Exception as e:
                                    logging.error(f"Skipping equation {i} due to rendering error: {e}")
                                    continue
                            
                            if images:
                                self.copy_images(images, original_text=text, equations=equations)
                            else:
                                self.status_var.set("No valid images rendered")
                        else:
                            logging.info("No valid LaTeX equations found in clipboard content")
                            self.status_var.set("No equations found")
                        
                        processed_text = text
                    elif not text:
                        logging.info("Clipboard is empty or contains non-text data")
                        self.status_var.set("Clipboard empty")
                
                time.sleep(1)
            except Exception as e:
                logging.error(f"Error in clipboard monitoring loop: {e}")
                self.status_var.set("Error in monitoring")
                time.sleep(1)

if __name__ == "__main__":
    root = tk.Tk()
    app = LatexClipboardApp(root)
    root.mainloop()