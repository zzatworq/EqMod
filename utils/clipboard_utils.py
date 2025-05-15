import win32clipboard
import logging
import time
import re

def set_clipboard_html(html_content):
    """Set HTML content to the Windows clipboard."""
    try:
        # Validate HTML content
        if not html_content or not isinstance(html_content, str):
            raise ValueError("HTML content must be a non-empty string")

        # Register HTML clipboard format
        try:
            CF_HTML = win32clipboard.RegisterClipboardFormat("HTML Format")
        except Exception as e:
            logging.error(f"Failed to register HTML clipboard format: {e}")
            raise ValueError("Cannot register HTML clipboard format")

        # Prepare HTML clipboard format
        html_header = (
            "Version:0.9\r\n"
            "StartHTML:0000000105\r\n"
            "EndHTML:{:010d}\r\n"
            "StartFragment:0000000141\r\n"
            "EndFragment:{:010d}\r\n"
            "<html><body>\r\n"
            "<!--StartFragment-->{}<!--EndFragment-->\r\n"
            "</body></html>"
        )
        
        # Calculate positions
        fragment = html_content
        full_html = html_header.format(
            len(html_header) + len(fragment),
            len(html_header) + len(fragment) - len("<!--EndFragment-->\r\n</body></html>"),
            fragment
        )
        
        # Convert to bytes
        html_bytes = full_html.encode('utf-8')
        
        # Set clipboard data
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(CF_HTML, html_bytes)
            logging.info("Successfully set HTML to clipboard")
        finally:
            win32clipboard.CloseClipboard()
        
    except Exception as e:
        logging.error(f"Failed to set clipboard HTML: {str(e)}")
        raise Exception(f"Failed to set clipboard HTML: {str(e)}")

def get_clipboard_text():
    """Retrieve text from the Windows clipboard."""
    try:
        win32clipboard.OpenClipboard()
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
            text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            logging.info("Retrieved text from clipboard")
            return text
        else:
            logging.info("No text data available in clipboard")
            return None
    except Exception as e:
        logging.error(f"Failed to get clipboard text: {e}")
        return None
    finally:
        try:
            win32clipboard.CloseClipboard()
        except:
            pass

def validate_base64(data):
    """Validate if the input string is valid base64."""
    try:
        base64_pattern = re.compile(r'^[A-Za-z0-9+/=]+$')
        if not base64_pattern.match(data):
            return False
        import base64
        base64.b64decode(data, validate=True)
        return True
    except Exception as e:
        logging.error(f"Base64 validation failed: {e}")
        return False