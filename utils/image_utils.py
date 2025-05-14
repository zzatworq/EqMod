import matplotlib.pyplot as plt
import numpy as np
from PIL import Image
import io
import logging

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

def render_latex_to_image(latex_string, text_color, font_size, dpi):
    """Render a LaTeX string to an image with dynamic sizing."""
    logging.info(f"Processing LaTeX string: {latex_string}")
    try:
        # Scale font size based on DPI for consistent appearance
        scaled_font_size = font_size * (dpi / 100)  # Normalize to 100 DPI baseline
        # Use a larger initial figure size to accommodate varying content
        base_width = 12  # Increased to handle wider equations
        base_height = 3  # Increased to handle taller equations
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
                   bbox_inches='tight', pad_inches=0.05)  # Reduced padding
        plt.close(fig)
        buffer.seek(0)
        img = Image.open(buffer).convert("RGBA")
        bbox = img.getbbox()
        if bbox:
            left, top, right, bottom = bbox
            padding = max(5, dpi // 20)  # Minimum padding of 5 pixels
            # Crop to content with padding
            left = max(0, left - padding)
            top = max(0, top - padding)
            right = min(img.width, right + padding)
            bottom = min(img.height, bottom + padding)
            img = img.crop((left, top, right, bottom))
            # Remove fixed height resizing; keep natural size up to a max
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
            logging.warning(f"No valid bounding box for equation: {latex_string}")
            return None
        if is_image_empty(img):
            logging.warning(f"Generated empty image for equation: {latex_string}")
            return None
        return img
    except Exception as e:
        logging.error(f"Error rendering LaTeX string '{latex_string}': {e}")
        return None